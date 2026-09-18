import asyncio
import os
import time
import pyotp


async def resolver_2fa(page):
    """Obtém o segredo das variáveis de ambiente, gera o código TOTP e

    autentica no Conta Azul garantindo que não expire na borda da janela de 30s.
    """
    secret = os.getenv("CONTA_AZUL_TOTP_SECRET", "").replace(" ", "").strip()
    if not secret:
        raise ValueError("A secret CONTA_AZUL_TOTP_SECRET não foi configurada!")

    print("Verificando se o 2FA foi solicitado...")

    # Seletor para identificar a tela de 2FA
    seletor_input_2fa = (
        "input[type='text'], input[type='tel'], input[inputmode='numeric']"
    )

    try:
        await page.wait_for_selector(seletor_input_2fa, timeout=8000)
    except Exception:
        print("Tela de 2FA não apareceu. Prosseguindo...")
        return

    print("Tela de 2FA detectada. Gerando código...")

    # Se faltar menos de 5 segundos para virar a janela do TOTP, aguarda a nova janela
    tempo_restante = 30 - (int(time.time()) % 30)
    if tempo_restante < 5:
        print(
            f"Janela de TOTP expirando em {tempo_restante}s. Aguardando novo ciclo..."
        )
        await asyncio.sleep(tempo_restante + 1)

    totp = pyotp.TOTP(secret)
    codigo = totp.now()
    print(f"Código gerado: {codigo}")

    # Localiza os campos de preenchimento
    inputs = page.locator(seletor_input_2fa)
    qtd_inputs = await inputs.count()

    if qtd_inputs >= 6:
        # Caso o Conta Azul use 6 caixas individuais
        for i, digito in enumerate(codigo):
            await inputs.nth(i).fill(digito)
            await page.wait_for_timeout(50)
    else:
        # Caso seja campo único: foca e digita simulando teclado real
        await inputs.first.click()
        await page.keyboard.type(codigo, delay=100)

    # Dispara o Enter e/ou clica no botão de submeter
    await page.keyboard.press("Enter")

    botao_submit = page.locator(
        "button[type='submit'], button:has-text('Continuar'), button:has-text('Entrar'), button:has-text('Autenticar')"
    ).first
    if await botao_submit.is_visible():
        try:
            await botao_submit.click(timeout=3000)
        except Exception:
            pass

    # Validação obrigatória: aguarda redirecionamento para fora da tela de autenticação
    try:
        await page.wait_for_url(
            lambda url: "login" not in url and "auth" not in url, timeout=20000
        )
        print("2FA autenticado com sucesso!")
    except Exception:
        await page.screenshot(path="erro_2fa.png")
        # IMPORTANTE: levantar exceção para NÃO seguir adiante deslogado
        raise RuntimeError(
            "Erro no fluxo do 2FA: 2FA preenchido, mas o login não foi autenticado."
        )
