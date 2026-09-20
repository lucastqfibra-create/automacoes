import asyncio
import datetime
import os
import subprocess
import sys
import time
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
from playwright.async_api import async_playwright

# Garantir que openpyxl esteja instalado para leitura de .xlsx
try:
  import openpyxl
except ImportError:
  print("Instalando openpyxl dinamicamente...")
  subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
  import openpyxl

# Carregar variáveis de ambiente
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
  """Carrega a planilha baixada do Conta Azul (suportando .xlsx, .xls ou .csv)

  e ajusta automaticamente o cabeçalho real.
  """
  with open(caminho_arquivo, "rb") as f:
    cabecalho_bytes = f.read(2048)

  # 1. Se for XLSX (ZIP container: PK\x03\x04)
  if cabecalho_bytes.startswith(b"PK\x03\x04"):
    print("Formato detectado: Excel .xlsx")
    try:
      df = pd.read_excel(caminho_arquivo, engine="openpyxl")
    except Exception:
      df = pd.read_excel(caminho_arquivo)

    # Verifica se as colunas estão na primeira linha ou se há cabeçalho antes
    colunas_str = " ".join([str(c).upper() for c in df.columns])
    if "CFOP" not in colunas_str and "NFE" not in colunas_str:
      for idx, row in df.head(10).iterrows():
        row_str = " ".join([str(val).upper() for val in row.values])
        if "CFOP" in row_str or "NFE" in row_str or "NÚMERO" in row_str:
          df.columns = df.iloc[idx]
          df = df.iloc[idx + 1 :].reset_index(drop=True)
          break
    return df

  # 2. Se for XLS binário antigo
  if cabecalho_bytes.startswith(b"\xd0\xcf\x11\xe0"):
    print("Formato detectado: Excel .xls binário")
    try:
      return pd.read_excel(caminho_arquivo)
    except Exception:
      pass

  # 3. Se for tabela HTML
  if b"<html" in cabecalho_bytes.lower() or b"<table" in cabecalho_bytes.lower():
    print("Formato detectado: Tabela HTML")
    dfs = pd.read_html(caminho_arquivo)
    if dfs:
      return dfs[0]

  # 4. Tratar como CSV / Texto com cabeçalho variável
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
            print(
                f"CSV lido com sucesso! (sep='{sep}', skiprows={linha_cabecalho},"
                f" encoding='{enc}')"
            )
            return df
        except Exception:
          continue
    except Exception:
      continue

  return pd.read_csv(
      caminho_arquivo, sep=None, engine="python", on_bad_lines="skip"
  )


async def processar_cliente(context, page, cliente_info, sheet):
  nome_cliente = cliente_info["nome"]
  celula_alvo = cliente_info["celula"]

  print(f"\n--- Iniciando processamento para: {nome_cliente} ---")

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
    print(f"Tentativa de navegação pelo menu: {e}")

  try:
    print(f"Selecionando o cliente {nome_cliente}...")
    client_row = page.locator("tr", has_text=nome_cliente).first
    await client_row.wait_for(state="visible", timeout=25000)
    await client_row.click(force=True)
    await asyncio.sleep(2)

    # Clicar em Acessar CA Pro
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
      print(f"Aviso no clique do CA Pro: {e}")

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
      screenshot_path = f"erro_{nome_cliente.replace(' ', '_')}.png"
      await page.screenshot(path=screenshot_path)
      print(f"Screenshot de erro salvo em: {screenshot_path}")
    except Exception:
      pass
    return

  download_path = None

  # --- Navegação para Vendas -> NF-e ---
  try:
    print("Aguardando carregamento da interface do CA Pro...")
    seletor_menu_vendas = (
        '#PRODUCTS, [id*="PRODUCTS"], :has-text("Vendas"),'
        ' :has-text("Produtos")'
    )
    menu_vendas = page_pro.locator(seletor_menu_vendas).first
    await menu_vendas.wait_for(state="visible", timeout=25000)
    await asyncio.sleep(1)

    print("Navegando pelo menu Vendas -> NF-e...")
    await menu_vendas.click(force=True)
    await asyncio.sleep(1.5)

    seletor_menu_nfe = (
        '#SALES_CONTROL_PRODUCT_INVOICE, :has-text("Notas fiscais de produto"),'
        ' :has-text("Notas fiscais")'
    )
    menu_nfe = page_pro.locator(seletor_menu_nfe).first
    await menu_nfe.wait_for(state="visible", timeout=15000)
    await menu_nfe.click(force=True)
    print("Acessou a tela de Notas Fiscais de Produto!")

    await page_pro.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(4)

    # --- Configurar filtro para "Este mês" ---
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
          print("Filtro alterado para 'Este mês' com sucesso!")

      await page_pro.wait_for_load_state("domcontentloaded")
      await asyncio.sleep(3)

      try:
        lupa2 = (
            page_pro.locator('input[placeholder*="Pesquisar"]')
            .locator("xpath=..")
            .locator("button")
            .first
        )
        if await lupa2.count() > 0 and await lupa2.is_visible():
          await lupa2.click()
          await page_pro.wait_for_load_state("domcontentloaded")
          await asyncio.sleep(3)
      except Exception:
        pass

    except Exception as e:
      print(f"Aviso no filtro de data: {e}")

    # --- Exportação da Planilha de NF-e ---
    # 1. Marcar checkbox do cabeçalho da tabela (selecionar todas as notas)
    try:
      chk_todos = page_pro.locator(
          'th input[type="checkbox"], thead input[type="checkbox"]'
      ).first
      if await chk_todos.count() > 0 and await chk_todos.is_visible():
        if not await chk_todos.is_checked():
          await chk_todos.click(force=True)
          print("Checkbox de selecionar todas as notas marcado com sucesso!")
          await asyncio.sleep(1)
    except Exception as e:
      print(f"Aviso no checkbox da tabela: {e}")

    # 2. Clicar no botão 'Ações' da barra superior (excluindo linhas de notas)
    print("Abrindo menu de exportação (Ações no cabeçalho)...")
    clicou_menu = False
    btn_acoes_header = (
        page_pro.locator(
            'button:has-text("Ações"), button:has-text("Ações em lote"),'
            ' [aria-label*="Ações"], div[title="Ações"] button,'
            ' button:has-text("Exportar")'
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

    # 3. Clicar na opção Exportar planilha
    print("Clicando na opção Exportar planilha...")
    opcao_exportar = page_pro.locator(
        ':has-text("Exportar planilha"), [role="menuitem"]:has-text("Exportar"),'
        ' a:has-text("Exportar"), li:has-text("Exportar"), span:has-text("Exportar")'
    ).last

    await opcao_exportar.wait_for(state="visible", timeout=15000)

    async with page_pro.expect_download(timeout=45000) as download_info:
      try:
        await opcao_exportar.click(timeout=5000)
      except Exception:
        await opcao_exportar.evaluate("el => el.click()")

    download = await download_info.value
    download_path = await download.path()
    print(f"Planilha baixada em: {download_path}")

  except Exception as e:
    print(f"Erro durante a navegação na Conta Azul ({nome_cliente}): {e}")
    try:
      screenshot_path = f"erro_{nome_cliente.replace(' ', '_')}.png"
      await page_pro.screenshot(path=screenshot_path)
      print(f"Screenshot de erro salvo em: {screenshot_path}")
    except Exception:
      pass
    await page_pro.close()
    return

  await page_pro.close()

  # --- Leitura e Cálculo dos Dados Baixados ---
  if download_path:
    print("Processando dados do arquivo baixado...")
    df = carregar_planilha(download_path)

    # Identificar colunas dinamicamente
    col_cfop = next(
        (c for c in df.columns if "CFOP" in str(c).upper()), "CFOP"
    )
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
    col_total = next(
        (
            c
            for c in df.columns
            if "TOTAL" in str(c).upper() or "VALOR" in str(c).upper()
        ),
        "Total NF-e",
    )

    print(
        f"Colunas mapeadas: CFOP='{col_cfop}', Número='{col_numero}',"
        f" Total='{col_total}'"
    )

    if "Fibrart" in nome_cliente:
      cfops_validos = ["5101", "6101"]
    elif "Afonso" in nome_cliente:
      cfops_validos = ["5102", "6102"]
    else:
      cfops_validos = ["5101"]

    padrao_regex = "|".join(cfops_validos)

    if col_cfop in df.columns:
      df[col_cfop] = df[col_cfop].astype(str)
      df_filtrado = df[df[col_cfop].str.contains(padrao_regex, na=False)]
    else:
      df_filtrado = df

    if col_numero in df_filtrado.columns:
      df_unique_nfe = df_filtrado.drop_duplicates(subset=[col_numero])
    else:
      df_unique_nfe = df_filtrado

    def parse_money(valor_str):
      if pd.isna(valor_str):
        return 0.0
      valor = (
          str(valor_str)
          .replace("R$", "")
          .replace(".", "")
          .replace(",", ".")
          .strip()
      )
      try:
        return float(valor)
      except ValueError:
        return 0.0

    if df_unique_nfe.empty or col_total not in df_unique_nfe.columns:
      total_cfop = 0.0
    else:
      df_unique_nfe["Total NF-e Num"] = df_unique_nfe[col_total].apply(
          parse_money
      )
      total_cfop = df_unique_nfe["Total NF-e Num"].sum()
      if isinstance(total_cfop, str):
        total_cfop = 0.0
  else:
    total_cfop = 0.0

  total_formatado = (
      f"{float(total_cfop):,.2f}"
      .replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  print(f"Total Calculado para {nome_cliente}: R$ {total_formatado}")

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
        print(f"Login direto realizado com sucesso! URL: {url_atual}")
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
