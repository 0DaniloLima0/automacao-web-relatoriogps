import os
import shutil
import openpyxl
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

# --- CONFIGURAÇÕES DE CAMINHO ---
# Usamos variáveis de ambiente ou caminhos genéricos para não expor a estrutura da rede
PASTA_DESTINO = os.getenv("CAMINHO_PRODUCAO", r"DIRETÓRIO")
PASTA_OLD = os.getenv("CAMINHO_BACKUP", r"DIRETÓRIO")
NOME_FINAL = "Base GPS.xlsx"

def gerenciar_arquivos_antigos():
    """Move o arquivo atual para a pasta de backup antes de baixar o novo."""
    arquivo_atual = os.path.join(PASTA_DESTINO, NOME_FINAL)
    
    if os.path.exists(arquivo_atual):
        if not os.path.exists(PASTA_OLD):
            os.makedirs(PASTA_OLD)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_backup = f"Base_GPS_OLD_{timestamp}.xlsx"
        
        print(f"Movendo arquivo antigo para: {nome_backup}")
        shutil.move(arquivo_atual, os.path.join(PASTA_OLD, nome_backup))

def tratar_planilha(caminho_arquivo):
    """Renomeia a aba (sheet) interna para 'base'."""
    try:
        print("Ajustando nome da aba para 'base'...")
        wb = openpyxl.load_workbook(caminho_arquivo)
        ws = wb.active
        ws.title = "base"
        wb.save(caminho_arquivo)
        wb.close()
    except Exception as e:
        print(f"Erro ao tratar planilha: {e}")

def run(playwright: Playwright) -> None:
    # Lança o navegador
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    # --- CREDENCIAIS E URL (Extraídas do ambiente para segurança) ---
    URL_SISTEMA = os.getenv("URL_SISTEMA", "https://portal.exemplo.com.br/login")
    USER_SISTEMA = os.getenv("USER_SISTEMA", "seu_usuario@email.com")
    PASS_SISTEMA = os.getenv("PASS_SISTEMA", "sua_senha_secreta")

    print("Iniciando acesso ao portal...")
    page.goto(URL_SISTEMA)

    # Login
    page.get_by_role("textbox", name="Email").fill(USER_SISTEMA)
    page.get_by_role("textbox", name="Password").fill(PASS_SISTEMA)
    page.get_by_role("button", name="Sign in").click()

    # Aguarda o carregamento e clica para exportar
    print("Aguardando carregamento da página principal...")
    page.wait_for_load_state("networkidle")
    
    # Navegação nos menus
    page.get_by_role("button", name="Opções").click()

    # Captura o evento de download
    with page.expect_download() as download_info:
        page.get_by_role("menuitem", name="Exportar Planilha").click()
    
    download = download_info.value

    # --- PROCESSAMENTO DE ARQUIVOS ---
    gerenciar_arquivos_antigos()

    # Salva o novo arquivo
    caminho_completo = os.path.join(PASTA_DESTINO, NOME_FINAL)
    download.save_as(caminho_completo)
    print(f"Download concluído: {caminho_completo}")

    # Tratamento da planilha
    tratar_planilha(caminho_completo)

    print("Processo finalizado com sucesso!")

    # Fecha o navegador
    context.close()
    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)