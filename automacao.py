from playwright.sync_api import sync_playwright
from openpyxl import load_workbook

# ==========================================
# CONFIGURAÇÕES
# ==========================================

ARQUIVO_EXCEL = "Levantamento semanal de multas.xlsx"
ABA = "AUTOPRF"

URL_PRF = "https://pesquisa-auto.prf.gov.br/#/pesquisa/consultar-debitos"


# ==========================================
# LER PLANILHA
# ==========================================

workbook = load_workbook(ARQUIVO_EXCEL)
planilha = workbook[ABA]

linha = 2

placa = planilha.cell(linha, 2).value
renavam = planilha.cell(linha, 4).value


print("-----------------------------------")
print("VEÍCULO ENCONTRADO NA PLANILHA")
print("-----------------------------------")
print(f"Placa: {placa}")
print(f"RENAVAM: {renavam}")
print("-----------------------------------")


# ==========================================
# ABRIR NAVEGADOR
# ==========================================

with sync_playwright() as p:

    print("Abrindo navegador...")

    browser = p.chromium.launch(
        headless=False
    )

    page = browser.new_page()

    print("Abrindo site da PRF...")

    page.goto(URL_PRF)

    page.wait_for_timeout(5000)

    print("Site da PRF aberto.")


    # ==========================================
    # PREENCHER PLACA
    # ==========================================

    print("Preenchendo placa...")

    page.get_by_label("Placa").fill(str(placa))

    print("Placa preenchida.")


    # ==========================================
    # PREENCHER RENAVAM
    # ==========================================

    print("Preenchendo RENAVAM...")

    page.get_by_label("RENAVAM").fill(str(renavam))

    print("RENAVAM preenchido.")


    # ==========================================
    # PESQUISAR
    # ==========================================

    print("Clicando em Pesquisar...")

    page.get_by_role(
        "button",
        name="Pesquisar"
    ).click()

    print("Consulta realizada.")


    # ==========================================
    # AGUARDAR RESULTADO
    # ==========================================

    print("Aguardando resultado...")

    page.wait_for_timeout(10000)

    print("Resultado carregado.")

    input("\nPressione ENTER para fechar o navegador...")

    browser.close()