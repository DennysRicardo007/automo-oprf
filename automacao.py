from playwright.sync_api import sync_playwright
from openpyxl import load_workbook


# ==========================================
# CONFIGURAÇÕES
# ==========================================

ARQUIVO_EXCEL = "Levantamento Semanal de Multas.xlsx"
ABA = "AUTOPRF"

URL_PRF = "https://pesquisa-auto.prf.gov.br/#/pesquisa/consultar-debitos"


# ==========================================
# LER A PLANILHA
# ==========================================

workbook = load_workbook(ARQUIVO_EXCEL)

planilha = workbook[ABA]

# Primeira linha de veículos = linha 2
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
# ABRIR O NAVEGADOR
# ==========================================

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    page = browser.new_page()

    # Abre o AutoPRF
    page.goto(URL_PRF)

    # Espera a página carregar
    page.wait_for_timeout(5000)

    print("Site da PRF aberto.")


    # ======================================
    # PREENCHER PLACA
    # ======================================

    # ATENÇÃO:
    # Vamos ajustar esses seletores depois
    # de confirmar como o site identifica
    # os campos.

    page.get_by_label("Placa").fill(str(placa))

    print("Placa preenchida.")


    # ======================================
    # PREENCHER RENAVAM
    # ======================================

    page.get_by_label("RENAVAM").fill(str(renavam))

    print("RENAVAM preenchido.")


    # ======================================
    # CLICAR EM PESQUISAR
    # ======================================

    page.get_by_role(
        "button",
        name="Pesquisar"
    ).click()

    print("Consulta realizada.")

    # Deixa a página aberta para vermos o resultado
    page.wait_for_timeout(10000)

    print("Resultado carregado.")

    # Não fecha o navegador ainda
    input("Pressione ENTER para fechar o navegador...")

    browser.close()