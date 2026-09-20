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

  for enc in ["utf-8", "latin1", "cp1252"]:
    for sep in [";", ",", "\t"]:
      try:
        df = pd.read_csv(
            caminho_arquivo, sep=sep, encoding=enc, on_bad_lines="skip"
        )
        if len(df.columns) > 3:
          return df
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
  await asyncio.sleep(3)

  # Espera as linhas da tabela de clientes carregarem
  try:
    await page.wait_for_selector(
        "table tbody tr, tr[class*='row']", timeout=25000
    )
  except Exception:
    pass

  try:
    print(f"Selecionando o cliente {nome_cliente}...")
    client_row = page.locator("tr", has_text=nome_cliente).first
    await client_row.wait_for(state="visible", timeout=30000)
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
    print("Aguardando carregamento da interface do CA Pro...")
    await page_pro.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(3)

    # 1. Navegação explícita e obrigatória: Vendas -> Notas fiscais de produto
    print("Aguardando menu Vendas...")
    seletor_menu_vendas = (
        '#PRODUCTS, [id*="PRODUCTS"], nav a:has-text("Vendas"),'
        ' a:has-text("Vendas"), button:has-text("Vendas")'
    )
    menu_vendas = page_pro.locator(seletor_menu_vendas).first
    await menu_vendas.wait_for(state="visible", timeout=25000)
    await menu_vendas.click(force=True)
    await asyncio.sleep(1.5)

    print("Clicando em Notas fiscais de produto...")
    seletor_menu_nfe = (
        '#SALES_CONTROL_PRODUCT_INVOICE, a:has-text("Notas fiscais de'
        ' produto"), :has-text("Notas fiscais de produto"), a[href*="nota"]'
    )
    menu_nfe = page_pro.locator(seletor_menu_nfe).first
    await menu_nfe.wait_for(state="visible", timeout=20000)
    await menu_nfe.click(force=True)
    print("Acessou a tela de Notas Fiscais de Produto com sucesso!")

    await page_pro.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(4)

    # 2. Configuração do filtro para 'Este mês'
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
    except Exception as e:
      print(f"Aviso no filtro de data: {e}")

    # 3. Marcar checkbox do cabeçalho da tabela (selecionar todas as notas)
    print("Selecionando todas as notas da tabela...")
    chk_todos = page_pro.locator(
        'th input[type="checkbox"], thead input[type="checkbox"]'
    ).first
    if await chk_todos.count() > 0:
      if not await chk_todos.is_checked():
        await chk_todos.click(force=True)
      print("Checkbox do cabeçalho marcado com sucesso!")
      await asyncio.sleep(1.5)
    else:
      btn_sel = page_pro.locator('button:has-text("Selecionar todas")').first
      if await btn_sel.count() > 0 and await btn_sel.is_visible():
        await btn_sel.click(force=True)
        print("Botão 'Selecionar todas' clicado!")
        await asyncio.sleep(1.5)

    # 4. Clicar no botão 'Ações' do cabeçalho (fora das linhas individuais)
    print("Abrindo menu de exportação do cabeçalho...")
    btn_header = (
        page_pro.locator(
            'button:has-text("Ações em lote"), button:has-text("Ações"),'
            ' button:has-text("Exportar"), [aria-label*="Ações"]'
        )
        .filter(has_not=page_pro.locator("tbody tr *"))
        .first
    )

    if await btn_header.count() > 0:
      await btn_header.click(force=True)
      await asyncio.sleep(1.5)

    # 5. Clicar em 'Exportar planilha'
    print("Clicando em Exportar planilha...")
    opcao_exportar = page_pro.locator(
        ':has-text("Exportar planilha"), [role="menuitem"]:has-text("Exportar"),'
        ' a:has-text("Exportar"), li:has-text("Exportar"),'
        ' span:has-text("Exportar"), button:has-text("Exportar")'
    ).last

    await opcao_exportar.wait_for(state="visible", timeout=20000)

    async with page_pro.expect_download(timeout=45000) as download_info:
      try:
        await opcao_exportar.click(timeout=5000)
      except Exception:
        await opcao_exportar.evaluate("el => el.click()")

    download = await download_info.value
    download_path = await download.path()
    print(f"Planilha de NF-e baixada em: {download_path}")

  except Exception as e:
    print(f"Erro na navegação ({nome_cliente}): {e}")
    try:
      await page_pro.screenshot(path=f"erro_{nome_cliente.replace(' ', '_')}.png")
    except Exception:
      pass
    await page_pro.close()
    return

  await page_pro.close()

  # Processamento e Cálculo dos Valores
  total_calculado = 0.0
  if download_path:
    print("Processando dados do arquivo baixado...")
    df = carregar_planilha(download_path)

    print(f"Total de linhas na planilha: {len(df)}")
    print(f"Colunas encontradas: {list(df.columns)}")

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

    # Identificar a coluna real do Valor Total da Nota Fiscal (ignorando parcelas)
    col_total = None
    for c in df.columns:
      c_up = str(c).upper()
      if "VALOR TOTAL" in c_up and ("NOTA" in c_up or "NF" in c_up):
        col_total = c
        break
    if not col_total:
      for c in df.columns:
        c_up = str(c).upper()
        if "VALOR TOTAL" in c_up and "PRODUTO" in c_up:
          col_total = c
          break
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
    if not col_total:
      col_total = next(
          (
              c
              for c in df.columns
              if "TOTAL" in str(c).upper() and "PARCELA" not in str(c).upper()
          ),
          df.columns[-1],
      )

    print(
        f"Colunas selecionadas: CFOP='{col_cfop}', Número='{col_numero}',"
        f" ValorTotal='{col_total}'"
    )

    if "Fibrart" in nome_cliente:
      cfops_alvo = ["5101", "6101"]
    elif "Afonso" in nome_cliente:
      cfops_alvo = ["5102", "6102"]
    else:
      cfops_alvo = ["5101"]

    if col_cfop in df.columns:
      # Extrai os primeiros 4 dígitos (atende 5.101 e 5101)
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

    if col_numero in df_filtrado.columns:
      df_unique_nfe = df_filtrado.drop_duplicates(subset=[col_numero])
    else:
      df_unique_nfe = df_filtrado

    if not df_unique_nfe.empty and col_total and col_total in df_unique_nfe.columns:
      df_unique_nfe["Total NF-e Num"] = df_unique_nfe[col_total].apply(
          parse_money
      )
      total_calculado = df_unique_nfe["Total NF-e Num"].sum()
      if isinstance(total_calculado, str):
        total_calculado = 0.0

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
    await page.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(2)

    # 2FA
    if CONTA_AZUL_TOTP_SECRET:
      try:
        print("Verificando se o 2FA foi solicitado...")
        await asyncio.sleep(2)
        seletor_inputs = (
            'input:visible:not([type="checkbox"]):not([type="radio"])'
        )

        try:
          await page.wait_for_selector(seletor_inputs, timeout=6000)
        except Exception:
          pass

        inputs_2fa = page.locator(seletor_inputs)
        qtd_inputs = await inputs_2fa.count()

        is_2fa = False
        if qtd_inputs > 0 and ("login" in page.url or "auth" in page.url):
          primeiro_tipo = (
              await inputs_2fa.first.get_attribute("type") or ""
          ).lower()
          primeiro_nome = (
              await inputs_2fa.first.get_attribute("name") or ""
          ).lower()
          if primeiro_tipo != "email" and "email" not in primeiro_nome:
            is_2fa = True

        if is_2fa:
          print("Tela de 2FA detectada. Gerando código...")
          import pyotp

          tempo_restante = 30 - (int(time.time()) % 30)
          if tempo_restante < 5:
            await asyncio.sleep(tempo_restante + 1)

          secret_limpo = (
              CONTA_AZUL_TOTP_SECRET.replace(" ", "").strip().upper()
          )
          totp = pyotp.TOTP(secret_limpo)
          codigo_2fa = totp.now()
          print(f"Código gerado: {codigo_2fa}")

          if qtd_inputs >= 6:
            for i, digito in enumerate(codigo_2fa):
              await inputs_2fa.nth(i).fill(digito)
              await asyncio.sleep(0.05)
          else:
            await inputs_2fa.first.click()
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

          # Aguarda sair da tela de login
          for _ in range(12):
            await asyncio.sleep(1)
            if "login" not in page.url and "auth" not in page.url:
              break

          print("2FA preenchido com sucesso!")

      except Exception as e:
        print(f"Aviso no fluxo de 2FA: {e}")

    await page.wait_for_load_state("domcontentloaded")
    await asyncio.sleep(3)

    for cliente in CLIENTES_ALVO:
      await processar_cliente(context, page, cliente, sheet)

    print("\nTodos os clientes foram processados com sucesso!")
    await browser.close()


if __name__ == "__main__":
  asyncio.run(main())
