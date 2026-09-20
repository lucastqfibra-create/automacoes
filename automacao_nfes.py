import asyncio
import datetime
import os
import sys
import time
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
from playwright.async_api import async_playwright

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


async def processar_cliente(context, page, cliente_info, sheet):
  nome_cliente = cliente_info["nome"]
  celula_alvo = cliente_info["celula"]

  print(f"\n--- Iniciando processamento para: {nome_cliente} ---")

  # Voltar para a tela de clientes
  print("Acessando lista de clientes no Conta Azul Mais...")
  await page.goto("https://mais.contaazul.com/#/inicio")
  await page.wait_for_load_state("networkidle")
  await asyncio.sleep(2)

  try:
    menu_locator = page.locator('text="Meus clientes"').first
    await menu_locator.evaluate("el => el.click()")
  except Exception as e:
    print("Tentando fallback de clique em Meus Clientes...")
    await page.click('text="Meus clientes"', force=True)

  await page.wait_for_load_state("networkidle")
  await asyncio.sleep(3)

  # Espera a lista carregar para não dar timeout na Fibrart
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

    # Clicar em Acessar CA Pro que está visível
    btn_pro_els = page.locator('text="Acessar CA Pro"')
    page_pro = None
    for i in range(await btn_pro_els.count()):
      if await btn_pro_els.nth(i).is_visible():
        try:
          # Inicia a escuta pelo evento da nova aba em background
          page_promise = asyncio.create_task(
              context.wait_for_event("page", timeout=15000)
          )

          await btn_pro_els.nth(i).click(force=True)
          await asyncio.sleep(2)

          # Verifica se o modal "Existe uma sessão ativa" apareceu
          modal_confirmar = page.locator('button:has-text("Confirmar")').first
          if (
              await modal_confirmar.count() > 0
              and await modal_confirmar.is_visible()
          ):
            print("Modal de sessão ativa detectado. Clicando em Confirmar...")
            await modal_confirmar.click(force=True)

          page_pro = await page_promise
          break
        except Exception as e:
          print(f"Tentativa de clique {i} falhou: {e}")
          pass

    if not page_pro:
      await page.screenshot(
          path=f"erro_ca_pro_{nome_cliente.replace(' ', '_')}.png",
          full_page=True,
      )
      raise Exception(
          "Não abriu a aba do CA Pro após tentar os botões visíveis!"
      )

    await page_pro.wait_for_load_state("networkidle")
    print(f"Acessou CA Pro de {nome_cliente}.")

  except Exception as e:
    print(f"Erro ao selecionar cliente {nome_cliente}: {e}")
    try:
      screenshot_path = f"erro_{nome_cliente.replace(' ', '_')}.png"
      await page.screenshot(path=screenshot_path)
      print(f"Screenshot de erro salvo em: {screenshot_path}")
    except:
      pass
    return

  download_path = None
  # Navegando pelo menu Vendas -> NF-e (seu seletor original que funciona)
  try:
    vendas_els = page_pro.locator('//*[@id="PRODUCTS"]/div[1]/div')
    for i in range(await vendas_els.count()):
      if await vendas_els.nth(i).is_visible():
        await vendas_els.nth(i).click(force=True)
        break
    await asyncio.sleep(1)

    nfe_els = page_pro.locator('//*[@id="SALES_CONTROL_PRODUCT_INVOICE"]')
    for i in range(await nfe_els.count()):
      if await nfe_els.nth(i).is_visible():
        await nfe_els.nth(i).click(force=True)
        break
    await page_pro.wait_for_load_state("networkidle")
    await asyncio.sleep(3)  # Aguardar a lista renderizar inicialmente

    # Verifica se existem notas fiscais na tela
    vazio = False
    try:
      await page_pro.wait_for_selector(
          'text="Nenhum resultado encontrado"', timeout=15000
      )
      vazio = True
    except:
      vazio = False

    if vazio:
      print("Nenhuma nota fiscal encontrada no período inicial.")

    # Configurar filtro "Este mês"
    print("Tentando configurar filtro para 'Este mês'...")
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
      ]:
        btn = page_pro.locator(f'button:has-text("{texto}")').first
        if await btn.count() > 0 and await btn.is_visible():
          await btn.click(force=True)
          clicou_dropdown = True
          break

      if not clicou_dropdown:
        lbl_periodo = page_pro.locator('text="Período"').first
        if await lbl_periodo.count() > 0:
          pass

      if clicou_dropdown:
        await asyncio.sleep(1.5)
        btn_este_mes = page_pro.locator('text="Este mês"').nth(0)
        if await btn_este_mes.count() > 0 and await btn_este_mes.is_visible():
          await btn_este_mes.click()
          print("Filtro alterado para 'Este mês' com sucesso via Playwright!")
        else:
          btn_este_mes_alt = page_pro.locator('text="Este Mês"').nth(0)
          if (
              await btn_este_mes_alt.count() > 0
              and await btn_este_mes_alt.is_visible()
          ):
            await btn_este_mes_alt.click()
            print("Filtro alterado para 'Este Mês' com sucesso!")

        await page_pro.wait_for_load_state("networkidle")
        await asyncio.sleep(2)

        try:
          lupa2 = (
              page_pro.locator('input[placeholder*="Pesquisar"]')
              .locator("xpath=..")
              .locator("button")
              .first
          )
          if await lupa2.count() > 0 and await lupa2.is_visible():
            await lupa2.click()
            print("Clicou na lupa para forçar a busca.")
            await page_pro.wait_for_load_state("networkidle")
            await asyncio.sleep(3)
        except Exception as e:
          print(f"Não achou a lupa, mas o filtro já deve ter aplicado: {e}")

    except Exception as e:
      print(f"Não conseguiu alterar o filtro de data: {e}")

    # Re-verifica se ficou vazio após o filtro
    try:
      await page_pro.wait_for_selector(
          'text="Nenhum resultado encontrado"', timeout=5000
      )
      vazio = True
    except:
      vazio = False

    if vazio:
      print("Nenhuma nota fiscal encontrada no mês atual. O total será 0.")
    else:
      print("Abrindo menu de exportação (busca robusta)...")

      export_clicked = False
      for selector in [
          (
              'div[title="Ações"]'
              " .ds-split-button-wrapper-group__trigger button"
          ),
          'button:has-text("Exportar")',
          '[aria-label="Exportar"]',
          '[aria-label="Opções de exportação"]',
      ]:
        try:
          btn = page_pro.locator(selector).first
          if await btn.count() > 0 and await btn.is_visible():
            await btn.click(force=True)
            await asyncio.sleep(2)
            exportar_opt = page_pro.locator('text="Exportar planilha"').first
            if (
                await exportar_opt.count() > 0
                and await exportar_opt.is_visible()
            ):
              export_clicked = True
              break
        except:
          pass

      if not export_clicked:
        print(
            "Aviso: Não consegui confirmar a abertura do menu de exportação."
        )

      print("Clicando em Exportar planilha...")
      async with page_pro.expect_download() as download_info:
        btn_export = page_pro.locator('text="Exportar planilha"').first
        await btn_export.evaluate("el => el.click()")
      download = await download_info.value
      download_path = await download.path()
      print(f"Planilha baixada em: {download_path}")

  except Exception as e:
    print(f"Erro durante a navegação na Conta Azul ({nome_cliente}): {e}")
    try:
      screenshot_path = f"erro_{nome_cliente.replace(' ', '_')}.png"
      await page_pro.screenshot(path=screenshot_path)
      print(f"Screenshot de erro salvo em: {screenshot_path}")
    except:
      pass
    await page_pro.close()
    return

  await page_pro.close()

  if download_path:
    print("Processando dados da planilha baixada...")
    # Suporte a Excel (.xlsx) e CSV
    df = None
    try:
      with open(download_path, "rb") as f:
        cabecalho = f.read(4)
      if cabecalho.startswith(b"PK\x03\x04") or download_path.endswith(".xlsx"):
        df = pd.read_excel(download_path)
      else:
        df = pd.read_csv(download_path, sep=";", encoding="utf-8")
    except Exception:
      try:
        df = pd.read_csv(download_path, sep=";", encoding="latin1")
      except Exception:
        df = pd.read_excel(download_path)

    # Identificar coluna CFOP
    col_cfop = next((c for c in df.columns if "CFOP" in str(c).upper()), "CFOP")

    # Identificar coluna de número da nota
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

    # Identificar coluna de valor total (priorizando a nota e ignorando parcelas)
    col_total = None
    for c in df.columns:
      c_up = str(c).upper()
      if "VALOR TOTAL" in c_up and ("NOTA" in c_up or "NF" in c_up):
        col_total = c
        break
    if not col_total:
      for c in df.columns:
        c_up = str(c).upper()
        if "TOTAL" in c_up and not any(
            b in c_up for b in ["PARCELA", "ICMS", "IPI", "PIS", "COFINS"]
        ):
          col_total = c
          break
    if not col_total:
      col_total = "Total NF-e"

    print(
        f"Colunas utilizadas: CFOP='{col_cfop}', Número='{col_numero}',"
        f" Total='{col_total}'"
    )

    # Definir CFOPs aceitos por cliente (com ou sem ponto: 5101 e 5.101)
    if "Fibrart" in nome_cliente:
      cfops_base = ["5101", "6101"]
    elif "Afonso" in nome_cliente:
      cfops_base = ["5102", "6102"]
    else:
      cfops_base = ["5101"]

    cfops_expandidos = cfops_base + [c[0] + "." + c[1:] for c in cfops_base]
    padrao_regex = "|".join(cfops_expandidos)

    if col_cfop in df.columns:
      df_filtrado = df[
          df[col_cfop].astype(str).str.contains(padrao_regex, na=False)
      ]
    else:
      df_filtrado = df

    if col_numero in df_filtrado.columns:
      df_unique_nfe = df_filtrado.drop_duplicates(subset=[col_numero])
    else:
      df_unique_nfe = df_filtrado

    def parse_money(valor_str):
      if pd.isna(valor_str):
        return 0.0
      if isinstance(valor_str, (int, float)):
        return float(valor_str)
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
      total_calculado = 0.0
    else:
      df_unique_nfe["Total NF-e Num"] = df_unique_nfe[col_total].apply(
          parse_money
      )
      total_calculado = df_unique_nfe["Total NF-e Num"].sum()
      if isinstance(total_calculado, str):
        total_calculado = 0.0
  else:
    total_calculado = 0.0

  total_formatado = (
      f"{float(total_calculado):,.2f}"
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
        viewport={"width": 1041, "height": 947},
        locale="pt-BR",
        timezone_id="America/Sao_Paulo",
    )
    page = await context.new_page()

    print("Acessando Conta Azul...")
    await page.goto("https://mais.contaazul.com/#/login")

    await page.fill('input[type="email"]', CONTA_AZUL_EMAIL)
    await page.fill('input[type="password"]', CONTA_AZUL_PASSWORD)
    await page.click('text="Entrar"')
    await page.wait_for_load_state("networkidle")
    await asyncio.sleep(2)

    # --- Lógica de 2FA com PyOTP resiliente ---
    if CONTA_AZUL_TOTP_SECRET:
      try:
        print("Verificando se o 2FA foi solicitado...")
        try:
          await page.wait_for_selector(
              'text="aplicativo de autenticação"', timeout=5000
          )
        except:
          pass

        input_2fa_locator = page.locator(
            'input:visible:not([type="checkbox"])'
        ).first

        if await input_2fa_locator.count() > 0:
          print("Tela de 2FA detectada. Gerando código...")
          import pyotp

          tempo_restante = 30 - (int(time.time()) % 30)
          if tempo_restante < 5:
            await asyncio.sleep(tempo_restante + 1)

          totp = pyotp.TOTP(CONTA_AZUL_TOTP_SECRET.strip())
          codigo_2fa = totp.now()
          print(f"Código gerado: {codigo_2fa}")

          await input_2fa_locator.fill(codigo_2fa)
          await asyncio.sleep(1)

          btn_auth = page.locator(
              'button:has-text("Autenticar"):visible,'
              ' button:has-text("Confirmar"):visible,'
              ' button:has-text("Verificar"):visible'
          ).first
          if await btn_auth.count() > 0:
            await btn_auth.click()

          # Espera a transição de saída do login sem falso alarme
          for _ in range(12):
            await asyncio.sleep(1)
            if "login" not in page.url and not await btn_auth.is_visible():
              break

          print("2FA preenchido com sucesso!")
      except ImportError:
        print("ERRO FATAL: Biblioteca pyotp não instalada!")
      except Exception as e:
        print(f"Erro no fluxo do 2FA: {e}")
    else:
      print(
          "Aviso: CONTA_AZUL_TOTP_SECRET não configurado. Se pedir 2FA, vai"
          " falhar."
      )

    await page.wait_for_load_state("networkidle")
    await asyncio.sleep(3)

    # Processar cada cliente da lista
    for cliente in CLIENTES_ALVO:
      await processar_cliente(context, page, cliente, sheet)

    print("\nTodos os clientes foram processados com sucesso!")
    await browser.close()


if __name__ == "__main__":
  asyncio.run(main())
