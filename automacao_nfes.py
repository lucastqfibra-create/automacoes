import asyncio
import datetime
import os
import re
import subprocess
import sys
import time
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
from playwright.async_api import async_playwright

# Garantir openpyxl instalado
try:
  import openpyxl
except ImportError:
  print("Instalando openpyxl dinamicamente...")
  subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
  import openpyxl

load_dotenv()

CONTA_AZUL_EMAIL = os.getenv("CONTA_AZUL_EMAIL")
CONTA_AZUL_PASSWORD = os.getenv("CONTA_AZUL_PASSWORD")
GOOGLE_CRED_FILE = os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE")
SPREADSHEET_ID = os.getenv("GOOGLE_SPREADSHEET_ID")
CONTA_AZUL_TOTP_SECRET = os.getenv("CONTA_AZUL_TOTP_SECRET")

CLIENTES_ALVO = [
    {"nome": "Fibrart", "celula": "Q11"},
    {"nome": "Afonso Morais", "celula": "R11"},
]


def carregar_planilha(caminho_arquivo):
  """Carrega a planilha (.xlsx, .xls ou .csv) e ajusta o cabeçalho real."""
  with open(caminho_arquivo, "rb") as f:
    cabecalho_bytes = f.read(2048)

  if cabecalho_bytes.startswith(b"PK\x03\x04"):
    print("Formato detectado: Excel .xlsx")
    try:
      df = pd.read_excel(caminho_arquivo, engine="openpyxl")
    except Exception:
      df = pd.read_excel(caminho_arquivo)

    colunas_str = " ".join([str(c).upper() for c in df.columns])
    if "CFOP" not in colunas_str and "NFE" not in colunas_str:
      for idx, row in df.head(10).iterrows():
        row_str = " ".join([str(val).upper() for val in row.values])
        if "CFOP" in row_str or "NFE" in row_str or "NÚMERO" in row_str:
          df.columns = df.iloc[idx]
          df = df.iloc[idx + 1 :].reset_index(drop=True)
          break
    return df

  if cabecalho_bytes.startswith(b"\xd0\xcf\x11\xe0"):
    print("Formato detectado: Excel .xls binário")
    try:
      return pd.read_excel(caminho_arquivo)
    except Exception:
      pass

  if b"<html" in cabecalho_bytes.lower() or b"<table" in cabecalho_bytes.lower():
    print("Formato detectado: Tabela HTML")
    dfs = pd.read_html(caminho_arquivo)
    if dfs:
      return dfs[0]

  print("Processando arquivo como CSV/Texto...")
  for enc in ["utf-8", "latin1", "cp1252"]:
    try:
      with open(caminho_arquivo, "r", encoding=enc) as f:
        linhas = [f.readline() for _ in range(20)]

      linha_cabecalho = 0
      for idx, linha in enumerate(linhas):
        l_upper = linha.upper()
        if (
            "CFOP" in l_upper
            or "NÚMERO" in l_upper
            or "NUMERO" in l_upper
            or "NFE" in l_upper
        ):
          linha_cabecalho = idx
          break

      for sep in [";", ",", "\t"]:
        try:
          df = pd.read_csv(
              caminho_arquivo,
              sep=sep,
              encoding=enc,
              skiprows=linha_cabecalho,
              on_bad_lines="skip",
          )
          colunas_str = " ".join([str(c).upper() for c in df.columns])
          if (
              "CFOP" in colunas_str
              or "NFE" in colunas_str
              or "NUMERO" in colunas_str
          ):
            return df
        except Exception:
          continue
    except Exception:
      continue

  return pd.read_csv(
      caminho_arquivo, sep=None, engine="python", on_bad_lines="skip"
  )


def parse_money(valor):
  """Converte valores monetários para float de forma segura."""
  if pd.isna(valor):
    return 0.0
  if isinstance(valor, (int, float)):
    return float(valor)
  valor_str = str(valor).replace("R$", "").replace(" ", "").strip()
  if "," in valor_str:
    valor_str = valor_str.replace(".", "").replace(",", ".")
  try:
    return float(valor_str)
  except ValueError:
    return 0.0


async def processar_cliente(context, page, cliente_info, sheet):
  nome_cliente = cliente_info["nome"]
  celula_alvo = cliente_info["celula"]

  print(f"\n==========================================")
  print(f"Iniciando processamento para: {nome_cliente}")
  print(f"==========================================")

  print("Acessando lista de clientes no Conta Azul Mais...")
  await page.goto("https://mais.contaazul.com/#/clientes")
  await page.wait_for_load_state("domcontentloaded")
  await asyncio.sleep(2)

  try:
    if "clientes" not in page.url:
      menu_pai = page.locator(':has-text("Clientes")').first
      if await menu_pai.is_visible():
        await menu_pai.click()
        await asyncio.sleep(1)
      menu_meus_clientes = page.locator(':has-text("Meus clientes")').first
      await menu_meus_clientes.dispatch_event("click")
      await page.wait_for_load_state("domcontentloaded")
  except Exception as e:
    print(f"Aviso menu: {e}")

  try:
    print(f"Selecionando o cliente {nome_cliente}...")
    client_row = page.locator("tr", has_text=nome_cliente).first
    await client_row.wait_for(state="visible", timeout=25000)
    await client_row.click(force=True)
    await asyncio.sleep(2)

    print("Clicando em Acessar CA Pro...")
    btn_pro = page.locator('button:has-text("Acessar CA Pro")').first
    await btn_pro.wait_for(state="visible", timeout=20000)

    page_pro = None
    try:
      page_promise = asyncio.create_task(
          context.wait_for_event("page", timeout=20000)
      )
      await btn_pro.click(force=True)
      await asyncio.sleep(2)

      modal_confirmar = page.locator('button:has-text("Confirmar")').first
      if (
          await modal_confirmar.count() > 0
          and await modal_confirmar.is_visible()
      ):
        print("Modal de sessão ativa detectado. Clicando em Confirmar...")
        await modal_confirmar.click(force=True)

      page_pro = await page_promise
    except Exception as e:
      print(f"Aviso clique CA Pro: {e}")

    if not page_pro:
      await page.screenshot(
          path=f"erro_ca_pro_{nome_cliente.replace(' ', '_')}.png",
          full_page=True,
      )
      raise Exception("Não abriu a aba do CA Pro!")

    await page_pro.wait_for_load_state("domcontentloaded")
    print(f"Acessou CA Pro de {nome_cliente}.")

  except Exception as e:
    print(f"Erro ao selecionar cliente {nome_cliente}: {e}")
    try:
      await page.screenshot(path=f"erro_{nome_cliente.replace(' ', '_')}.png")
    except Exception:
      pass
    return

  download_path = None

  try:
    print("Aguardando interface do CA Pro...")
    seletor_menu_vendas = (
        '#PRODUCTS, [id*="PRODUCTS"], :has-text("Vendas"),'
        ' :has-text("Produtos")'
    )
    menu_vendas = page_pro.locator(seletor_menu_vendas).first
    await menu_vendas.wait_for(state="visible", timeout=25000)
    await asyncio.sleep(1)

    print("Navegando para Vendas -> NF-e...")
    await menu_vendas.click(force=True)
    await asyncio.sleep(1.5)

    seletor_menu_nfe = (
        '#SALES_CONTROL_PRODUCT_INVOICE, :has-text("Notas fiscais de produto"),'
        ' :has-text("Notas fiscais")'
    )
    menu_nfe = page_pro.locator(seletor_menu_nfe).first
    await menu_nfe.wait_for(state="visible", timeout=15000)
    await menu_nfe.click(force=True)
    print("Acessou tela de Notas Fiscais!")

    await page_pro.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(4)

    # Configurar filtro para "Este mês"
    print("Configurando filtro para 'Este mês'...")
    try:
      meses = [
          "Janeiro",
          "Fevereiro",
          "Março",
          "Abril",
          "Maio",
          "Junho",
          "Julho",
          "Agosto",
          "Setembro",
          "Outubro",
          "Novembro",
          "Dezembro",
      ]
      hoje = datetime.datetime.now()
      mes_atual_str = f"{meses[hoje.month - 1]} de {hoje.year}"

      clicou_dropdown = False
      for texto in [
          mes_atual_str,
          "Últimos 30 dias",
          "Mês passado",
          "Hoje",
          "Este ano",
          "Últimos 7 dias",
          "Período",
      ]:
        btn = page_pro.locator(f'button:has-text("{texto}")').first
        if await btn.count() > 0 and await btn.is_visible():
          await btn.click(force=True)
          clicou_dropdown = True
          break

      if clicou_dropdown:
        await asyncio.sleep(1.5)
        btn_este_mes = page_pro.locator(
            ':has-text("Este mês"), :has-text("Este Mês")'
        ).first
        if await btn_este_mes.count() > 0 and await btn_este_mes.is_visible():
          await btn_este_mes.click()
          print("Filtro alterado para 'Este mês'!")

      await page_pro.wait_for_load_state("domcontentloaded")
      await asyncio.sleep(3)
    except Exception as e:
      print(f"Aviso no filtro: {e}")

    # VERIFICAÇÃO PRÉVIA: Ver se há notas emitidas na tabela
    linhas_tabela = page_pro.locator(
        'table tbody tr:not([class*="empty"]):not([class*="no-data"])'
    )
    qtd_linhas = await linhas_tabela.count()
    texto_pagina = await page_pro.content()

    if (
        qtd_linhas == 0
        or "Nenhum registro" in texto_pagina
        or "Nenhuma nota" in texto_pagina
    ):
      print(
          f"Nenhuma nota fiscal emitida para {nome_cliente} neste período."
          " Definindo R$ 0,00 na planilha..."
      )
      sheet.update_acell(celula_alvo, "0,00")
      print(
          f"Planilha atualizada com sucesso na célula {celula_alvo}: R$ 0,00"
      )
      await page_pro.close()
      return

    # Marcar checkbox para selecionar todas as notas
    try:
      chk_todos = page_pro.locator(
          'th input[type="checkbox"], thead input[type="checkbox"]'
      ).first
      if await chk_todos.count() > 0 and await chk_todos.is_visible():
        if not await chk_todos.is_checked():
          await chk_todos.click(force=True)
          print("Checkbox de selecionar todas as notas marcado!")
          await asyncio.sleep(1)
    except Exception as e:
      print(f"Aviso no checkbox: {e}")

    # Clicar no botão 'Ações' do cabeçalho
    print("Abrindo menu Ações do cabeçalho...")
    clicou_menu = False
    btn_acoes_header = (
        page_pro.locator(
            'button:has-text("Ações"), button:has-text("Ações em lote"),'
            ' [aria-label*="Ações"], div[title="Ações"] button'
        )
        .filter(has_not=page_pro.locator("tbody tr *"))
        .first
    )

    if await btn_acoes_header.count() > 0 and await btn_acoes_header.is_visible():
      try:
        await btn_acoes_header.hover()
      except Exception:
        pass
      await btn_acoes_header.click(force=True)
      clicou_menu = True
      await asyncio.sleep(1.5)

    if not clicou_menu:
      await page_pro.evaluate("""() => {
          const els = Array.from(document.querySelectorAll('button, a, [role="button"]'))
              .filter(el => {
                  const t = (el.innerText || '').trim();
                  return (t === 'Ações' || t === 'Ações em lote' || t.includes('Exportar')) && !el.closest('tbody tr');
              });
          if (els.length > 0) els[0].click();
      }""")
      await asyncio.sleep(1.5)

    print("Clicando em Exportar planilha...")
    opcao_exportar = page_pro.locator(
        ':has-text("Exportar planilha"), [role="menuitem"]:has-text("Exportar"),'
        ' a:has-text("Exportar"), li:has-text("Exportar"), span:has-text("Exportar")'
    ).last

    await opcao_exportar.wait_for(state="visible", timeout=15000)

    try:
      async with page_pro.expect_download(timeout=35000) as download_info:
        try:
          await opcao_exportar.click(timeout=5000)
        except Exception:
          await opcao_exportar.evaluate("el => el.click()")

      download = await download_info.value
      download_path = await download.path()
      print(f"Planilha baixada em: {download_path}")
    except Exception as e:
      print(
          f"Download não gerado para {nome_cliente}: {e}. Assumindo sem notas."
      )
      download_path = None

  except Exception as e:
    print(f"Erro na navegação ({nome_cliente}): {e}")
    try:
      await page_pro.screenshot(path=f"erro_{nome_cliente.replace(' ', '_')}.png")
    except Exception:
      pass
    await page_pro.close()
    return

  await page_pro.close()

  # --- CÁLCULO PRECISO DOS VALORES ---
  total_calculado = 0.0

  if download_path:
    print("Analisando planilha baixada...")
    df = carregar_planilha(download_path)

    print(f"Total de linhas na planilha: {len(df)}")
    print(f"Todas as colunas encontradas: {list(df.columns)}")

    # 1. Identificar a coluna real do CFOP
    col_cfop = next(
        (c for c in df.columns if "CFOP" in str(c).upper()), "CFOP"
    )

    # 2. Identificar a coluna do Número da Nota
    col_numero = next(
        (
            c
            for c in df.columns
            if "NÚMERO" in str(c).upper()
            or "NUMERO" in str(c).upper()
            or "NFE" in str(c).upper()
        ),
        "Número da NFe",
    )

    # 3. IDENTIFICAÇÃO CORRETA DA COLUNA DE VALOR TOTAL (Ignorando parcelas)
    col_total = None
    # Prioridade 1: 'Valor total da nota fiscal'
    for c in df.columns:
      c_up = str(c).upper()
      if "VALOR TOTAL" in c_up and ("NOTA" in c_up or "NF" in c_up):
        col_total = c
        break
    # Prioridade 2: 'Valor total dos produtos'
    if not col_total:
      for c in df.columns:
        c_up = str(c).upper()
        if "VALOR TOTAL" in c_up and "PRODUTO" in c_up:
          col_total = c
          break
    # Prioridade 3: Qualquer coluna com 'VALOR' e 'NOTA' (sem parcelas nem impostos)
    if not col_total:
      for c in df.columns:
        c_up = str(c).upper()
        if ("TOTAL" in c_up or "VALOR" in c_up) and (
            "NOTA" in c_up or "NF" in c_up
        ):
          if not any(
              bad in c_up for bad in ["PARCELA", "ICMS", "IPI", "PIS", "COFINS"]
          ):
            col_total = c
            break

    print(
        f"Colunas selecionadas: CFOP='{col_cfop}', Número='{col_numero}',"
        f" ValorTotal='{col_total}'"
    )

    # Filtrar CFOPs com limpeza de pontuação (ex: 5.101 vira 5101)
    if "Fibrart" in nome_cliente:
      cfops_alvo = ["5101", "6101"]
    elif "Afonso" in nome_cliente:
      cfops_alvo = ["5102", "6102"]
    else:
      cfops_alvo = ["5101"]

    if col_cfop in df.columns:
      # Extrai os primeiros 4 dígitos numéricos de cada CFOP (funciona com 5.101, 5101, etc.)
      cfop_extraido = (
          df[col_cfop]
          .astype(str)
          .str.replace(".", "", regex=False)
          .str.extract(r"(\d{4})")[0]
      )
      print(
          f"CFOPs encontrados na planilha: {cfop_extraido.dropna().unique().tolist()}"
      )

      df_filtrado = df[cfop_extraido.isin(cfops_alvo)]
      print(
          f"Notas que atendem aos CFOPs {cfops_alvo}: {len(df_filtrado)} de"
          f" {len(df)}"
      )
    else:
      df_filtrado = df

    # Remover duplicadas pelo número da nota
    if col_numero in df_filtrado.columns:
      df_unique = df_filtrado.drop_duplicates(subset=[col_numero])
    else:
      df_unique = df_filtrado

    if not df_unique.empty and col_total and col_total in df_unique.columns:
      # Exibir amostra dos valores para auditoria no log
      amostra_valores = df_unique[col_total].head(5).tolist()
      print(f"Amostra de valores em '{col_total}': {amostra_valores}")

      df_unique["Valor_Numerico"] = df_unique[col_total].apply(parse_money)
      total_calculado = df_unique["Valor_Numerico"].sum()

  total_formatado = (
      f"{float(total_calculado):,.2f}"
      .replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  print(f"\n>>> TOTAL FINAL CALCULADO PARA {nome_cliente}: R$ {total_formatado}")

  print(f"Atualizando Google Sheets (Célula {celula_alvo})...")
  try:
    sheet.update_acell(celula_alvo, total_formatado)
    print(f"Planilha atualizada com sucesso na célula {celula_alvo}!")
  except Exception as e:
    print(f"Erro ao atualizar planilha para {nome_cliente}: {e}")


async def main():
  print("Iniciando automação múltipla...")

  scopes = [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
  ]
  credentials = Credentials.from_service_account_file(
      GOOGLE_CRED_FILE, scopes=scopes
  )
  client = gspread.authorize(credentials)
  sheet = client.open_by_key(SPREADSHEET_ID).sheet1

  async with async_playwright() as p:
    browser = await p.chromium.launch(headless=True)
    context = await browser.new_context(
        accept_downloads=True,
        viewport={"width": 1920, "height": 1080},
        locale="pt-BR",
        timezone_id="America/Sao_Paulo",
    )
    page = await context.new_page()

    print("Acessando Conta Azul...")
    await page.goto("https://mais.contaazul.com/#/login")
    await page.wait_for_load_state("domcontentloaded")

    await page.fill('input[type="email"]', CONTA_AZUL_EMAIL)
    await page.fill('input[type="password"]', CONTA_AZUL_PASSWORD)
    await page.click('text="Entrar"')

    print("Aguardando autenticação...")
    autenticado = False
    for _ in range(15):
      await asyncio.sleep(1)
      url_atual = page.url

      if "login" not in url_atual and "auth" not in url_atual:
        print(f"Login direto realizado! URL: {url_atual}")
        autenticado = True
        break

      inputs_visiveis = page.locator(
          'input:visible:not([type="checkbox"]):not([type="radio"])'
      )
      count = await inputs_visiveis.count()

      is_2fa = False
      if count > 0:
        primeiro_tipo = (
            await inputs_visiveis.first.get_attribute("type") or ""
        ).lower()
        primeiro_nome = (
            await inputs_visiveis.first.get_attribute("name") or ""
        ).lower()
        if primeiro_tipo != "email" and "email" not in primeiro_nome:
          is_2fa = True

      if is_2fa and CONTA_AZUL_TOTP_SECRET:
        print("Tela de 2FA detectada! Gerando código TOTP...")
        import pyotp

        tempo_restante = 30 - (int(time.time()) % 30)
        if tempo_restante < 5:
          await asyncio.sleep(tempo_restante + 1)

        secret_limpo = CONTA_AZUL_TOTP_SECRET.replace(" ", "").strip().upper()
        totp = pyotp.TOTP(secret_limpo)
        codigo_2fa = totp.now()
        print(f"Código 2FA gerado: {codigo_2fa}")

        if count >= 6:
          for i, digito in enumerate(codigo_2fa):
            await inputs_visiveis.nth(i).fill(digito)
            await asyncio.sleep(0.05)
        else:
          await inputs_visiveis.first.click()
          await page.keyboard.type(codigo_2fa, delay=100)

        await page.keyboard.press("Enter")
        await asyncio.sleep(1)

        btn_auth = page.locator(
            'button:has-text("Autenticar"):visible,'
            ' button:has-text("Confirmar"):visible,'
            ' button:has-text("Verificar"):visible,'
            ' button:has-text("Entrar"):visible'
        ).first
        if await btn_auth.count() > 0 and await btn_auth.is_visible():
          try:
            await btn_auth.click(timeout=3000)
          except Exception:
            pass

        for _ in range(12):
          await asyncio.sleep(1)
          if "login" not in page.url and "auth" not in page.url:
            autenticado = True
            break
        break

    if not autenticado and ("login" in page.url or "auth" in page.url):
      await page.screenshot(path="erro_login.png", full_page=True)
      raise RuntimeError(
          f"Falha na autenticação. A página permaneceu em: {page.url}"
      )

    print("Autenticação concluída com sucesso!")

    for cliente in CLIENTES_ALVO:
      await processar_cliente(context, page, cliente, sheet)

    print("\nTodos os clientes foram processados com sucesso!")
    await browser.close()


if __name__ == "__main__":
  asyncio.run(main())
