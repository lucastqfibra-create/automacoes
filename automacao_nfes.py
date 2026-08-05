import os
import asyncio
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

CONTA_AZUL_EMAIL = os.getenv('CONTA_AZUL_EMAIL')
CONTA_AZUL_PASSWORD = os.getenv('CONTA_AZUL_PASSWORD')
GOOGLE_CRED_FILE = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
SPREADSHEET_ID = os.getenv('GOOGLE_SPREADSHEET_ID')

CLIENTES_ALVO = [
    {"nome": "Fibrart", "celula": "Q11"},
    {"nome": "Afonso Morais", "celula": "R11"}
]

async def processar_cliente(context, page, cliente_info, sheet):
    nome_cliente = cliente_info["nome"]
    celula_alvo = cliente_info["celula"]
    
    print(f"\n--- Iniciando processamento para: {nome_cliente} ---")
    
    # Voltar para a tela de clientes
    print("Acessando lista de clientes no Conta Azul Mais...")
    await page.goto("https://mais.contaazul.com/#/inicio")
    await page.wait_for_load_state("networkidle")
    
    try:
        menu_locator = page.locator('text="Meus clientes"').first
        await menu_locator.evaluate("el => el.click()")
    except Exception as e:
        print("Tentando fallback de clique em Meus Clientes...")
        await page.click('text="Meus clientes"', force=True)
        
    await page.wait_for_load_state("networkidle")
    
    try:
        print(f"Selecionando o cliente {nome_cliente}...")
        client_row = page.locator('tr', has_text=nome_cliente).first
        await client_row.click(force=True)
        await asyncio.sleep(2)
        
        # Clicar em Acessar CA Pro que está visível
        btn_pro_els = page.locator('text="Acessar CA Pro"')
        page_pro = None
        for i in range(await btn_pro_els.count()):
            if await btn_pro_els.nth(i).is_visible():
                try:
                    # Inicia a escuta pelo evento da nova aba em background
                    page_promise = asyncio.create_task(context.wait_for_event('page', timeout=15000))
                    
                    await btn_pro_els.nth(i).click(force=True)
                    await asyncio.sleep(2)
                    
                    # Verifica se o modal "Existe uma sessão ativa" apareceu
                    modal_confirmar = page.locator('button:has-text("Confirmar")').first
                    if await modal_confirmar.count() > 0 and await modal_confirmar.is_visible():
                        print("Modal de sessão ativa detectado. Clicando em Confirmar...")
                        await modal_confirmar.click(force=True)
                        
                    page_pro = await page_promise
                    break
                except Exception as e:
                    print(f"Tentativa de clique {i} falhou: {e}")
                    pass
                
        if not page_pro:
            await page.screenshot(path=f"erro_ca_pro_{nome_cliente.replace(' ', '_')}.png", full_page=True)
            raise Exception("Não abriu a aba do CA Pro após tentar os botões visíveis!")
            
        await page_pro.wait_for_load_state("networkidle")
        print(f"Acessou CA Pro de {nome_cliente}.")
        
    except Exception as e:
        print(f"Erro ao selecionar cliente {nome_cliente}: {e}")
        return
    
    download_path = None
    # Navegando pelo menu Vendas -> NF-e
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
        await asyncio.sleep(3) # Aguardar a lista renderizar inicialmente
        
        # Verifica se existem notas fiscais na tela
        vazio = False
        try:
            # Aumentando para 15 segundos porque o React as vezes demora a popular o grid
            await page_pro.wait_for_selector('text="Nenhum resultado encontrado"', timeout=15000)
            vazio = True
        except:
            vazio = False
            
        if vazio:
            print("Nenhuma nota fiscal encontrada no período inicial.")
        
        # --- NOVO: Forçar o filtro "Este mês" (Puro Playwright) ---
        print("Tentando configurar filtro para 'Este mês'...")
        try:
            import datetime
            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            hoje = datetime.datetime.now()
            mes_atual_str = f"{meses[hoje.month - 1]} de {hoje.year}"
            
            clicou_dropdown = False
            
            # Tenta encontrar o botão pelo texto (ex: "Agosto de 2026" ou "Últimos 30 dias")
            for texto in [mes_atual_str, "Últimos 30 dias", "Mês passado", "Hoje", "Este ano", "Últimos 7 dias"]:
                btn = page_pro.locator(f'button:has-text("{texto}")').first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click(force=True)
                    clicou_dropdown = True
                    break
                    
            if not clicou_dropdown:
                # Estratégia B: Clicar abaixo de "Período"
                lbl_periodo = page_pro.locator('text="Período"').first
                if await lbl_periodo.count() > 0:
                    # Encontrar todos os botões na tela e clicar no 2º após o label (geralmente é <, [Mês], >)
                    pass # O loop acima costuma ser o suficiente
            
            if clicou_dropdown:
                await asyncio.sleep(1.5)
                # O menu abriu, agora clica em "Este mês"
                btn_este_mes = page_pro.locator('text="Este mês"').nth(0)
                if await btn_este_mes.count() > 0 and await btn_este_mes.is_visible():
                    await btn_este_mes.click() # Sem force=True para respeitar animações
                    print("Filtro alterado para 'Este mês' com sucesso via Playwright!")
                else:
                    btn_este_mes_alt = page_pro.locator('text="Este Mês"').nth(0)
                    if await btn_este_mes_alt.count() > 0 and await btn_este_mes_alt.is_visible():
                        await btn_este_mes_alt.click()
                        print("Filtro alterado para 'Este Mês' com sucesso!")
                
                await page_pro.wait_for_load_state("networkidle")
                await asyncio.sleep(2)
                
                # Forçar o refresh clicando na Lupa de pesquisa
                try:
                    lupa = page_pro.locator('button:has(svg), button[type="submit"]').filter(has_text="").nth(1)
                    # Melhor: localizar o botão ao lado do input de pesquisa
                    lupa2 = page_pro.locator('input[placeholder*="Pesquisar"]').locator('xpath=..').locator('button').first
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
            await page_pro.wait_for_selector('text="Nenhum resultado encontrado"', timeout=5000)
            vazio = True
        except:
            vazio = False

        if vazio:
            print("Nenhuma nota fiscal encontrada no mês atual. O total será 0.")
        else:
            print("Abrindo menu de exportação (busca robusta)...")
            
            export_clicked = False
            for selector in [
                'div[title="Ações"] .ds-split-button-wrapper-group__trigger button', # O botão da setinha ao lado de Ações
                'button:has-text("Exportar")',
                '[aria-label="Exportar"]',
                '[aria-label="Opções de exportação"]'
            ]:
                try:
                    btn = page_pro.locator(selector).first
                    if await btn.count() > 0 and await btn.is_visible():
                        await btn.click(force=True)
                        await asyncio.sleep(2)
                        # Tenta procurar a opção Exportar planilha no menu que se abriu
                        exportar_opt = page_pro.locator('text="Exportar planilha"').first
                        if await exportar_opt.count() > 0 and await exportar_opt.is_visible():
                            export_clicked = True
                            break
                except:
                    pass
            
            if not export_clicked:
                print("Aviso: Não consegui confirmar a abertura do menu de exportação.")
            
            print("Clicando em Exportar planilha...")
            async with page_pro.expect_download() as download_info:
                # Se não achar Exportar Planilha aqui, dará TimeoutError
                btn_export = page_pro.locator('text="Exportar planilha"').first
                await btn_export.evaluate("el => el.click()")
            download = await download_info.value
            download_path = await download.path()
            print(f"Planilha baixada em: {download_path}")
            
    except Exception as e:
        print(f"Erro durante a navegação na Conta Azul ({nome_cliente}): {e}")
        try:
            screenshot_path = f"C:\\Users\\lucas\\.gemini\\antigravity\\brain\\71932980-6b85-4838-96f3-898801f9af85\\erro_{nome_cliente.replace(' ', '_')}.png"
            await page_pro.screenshot(path=screenshot_path)
            print(f"Screenshot de erro salvo em: {screenshot_path}")
        except:
            pass
        await page_pro.close()
        return
        
    await page_pro.close()
    
    if download_path:
        print("Processando dados do CSV baixado...")
        try:
            df = pd.read_csv(download_path, sep=';', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(download_path, sep=';', encoding='latin1')
            
        # Definir CFOPs aceitos por cliente
        if "Fibrart" in nome_cliente:
            cfops_validos = ['5101', '6101']
        elif "Afonso" in nome_cliente:
            cfops_validos = ['5102', '6102']
        else:
            cfops_validos = ['5101']
            
        padrao_regex = '|'.join(cfops_validos)
        df_filtrado = df[df['CFOP'].str.contains(padrao_regex, na=False)]
        df_unique_nfe = df_filtrado.drop_duplicates(subset=['Número da NFe'])
        
        def parse_money(valor_str):
            if pd.isna(valor_str):
                return 0.0
            valor = str(valor_str).replace('R$', '').replace('.', '').replace(',', '.').strip()
            try:
                return float(valor)
            except ValueError:
                return 0.0
                
        if df_unique_nfe.empty:
            total_cfop_5101 = 0.0
        else:
            df_unique_nfe['Total NF-e Num'] = df_unique_nfe['Total NF-e'].apply(parse_money)
            total_cfop_5101 = df_unique_nfe['Total NF-e Num'].sum()
            if isinstance(total_cfop_5101, str):
                total_cfop_5101 = 0.0
    else:
        total_cfop_5101 = 0.0
    
    total_formatado = f"{float(total_cfop_5101):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    print(f"Total Calculado para {nome_cliente}: R$ {total_formatado}")
    
    print(f"Atualizando Google Sheets (Célula {celula_alvo})...")
    try:
        sheet.update_acell(celula_alvo, total_formatado)
        print(f"Planilha atualizada com sucesso na célula {celula_alvo}!")
    except Exception as e:
        print(f"Erro ao atualizar planilha para {nome_cliente}: {e}")

async def main():
    print("Iniciando automação múltipla...")
    
    # Preparar conexão com Google Sheets primeiro
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_file(GOOGLE_CRED_FILE, scopes=scopes)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            accept_downloads=True,
            viewport={'width': 1041, 'height': 947},
            locale='pt-BR',
            timezone_id='America/Sao_Paulo'
        )
        page = await context.new_page()
        
        print("Acessando Conta Azul...")
        await page.goto("https://mais.contaazul.com/#/login")
        
        await page.fill('input[type="email"]', CONTA_AZUL_EMAIL)
        await page.fill('input[type="password"]', CONTA_AZUL_PASSWORD)
        await page.click('text="Entrar"')
        await page.wait_for_load_state("networkidle")
        
        # Processar cada cliente da lista
        for cliente in CLIENTES_ALVO:
            await processar_cliente(context, page, cliente, sheet)
            
        print("\nTodos os clientes foram processados!")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
