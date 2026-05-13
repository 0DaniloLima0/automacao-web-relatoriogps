# automacao-web-relatoriogps
Automação para extração de relatório via web 

import os
import openpyxl
from playwright.sync_api import Playwright, sync_playwright

# --- CONFIGURAÇÕES GLOBAIS (Neutralizadas para GitHub) ---
NOME_ARQUIVO_FINAL = "dados_extraidos.xlsx"
NOME_ABA = "base"

# Credenciais e URLs via Variáveis de Ambiente
USUARIO = os.getenv("SISTEMA_USER", "seu_usuario_aqui")
SENHA = os.getenv("SISTEMA_PASS", "sua_senha_aqui")
URL_ALVO = os.getenv("SISTEMA_URL", "https://link-do-site.com.br/login")

def tratar_planilha(caminho_arquivo):
    """
    Abre o arquivo Excel baixado e renomeia a aba principal.
    """
    try:
        print(f"Limpando dados e ajustando aba para '{NOME_ABA}'...")
        wb = openpyxl.load_workbook(caminho_arquivo)
        ws = wb.active
        ws.title = NOME_ABA
        wb.save(caminho_arquivo)
        wb.close()
        print("✔ Tratamento de dados concluído com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao tratar planilha: {e}")

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    try:
        print("🚀 Iniciando automação...")
        page.goto(URL_ALVO)

        # Login genérico
        print("Realizando autenticação...")
        page.get_by_role("textbox", name="Email").fill(USUARIO)
        page.get_by_role("textbox", name="Password").fill(SENHA)
        page.get_by_role("button").click()

        page.wait_for_load_state("networkidle")
        
        # Ação de download (deve ser preenchida com o seletor específico do seu caso)
        with page.expect_download() as download_info:
            # page.click("#seu-botao-de-download") 
            pass 
        
        download = download_info.value
        caminho_completo = os.path.join(os.getcwd(), NOME_ARQUIVO_FINAL)
        download.save_as(caminho_completo)
        
        tratar_planilha(caminho_completo)
        print("✨ Fluxo de automação finalizado!")

    except Exception as e:
        print(f"❌ Ocorreu um erro inesperado: {e}")

    finally:
        print("🔒 Fechando navegador e limpando sessão...")
        context.close()
        browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
