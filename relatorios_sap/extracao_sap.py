from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import time
import tkinter as tk
from tkinter import simpledialog
from pathlib import Path
from datetime import date


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_RAIZ = Path(r"X:\Comum-PagamentosDiarios")

MESES = [
    "01 - JANEIRO",
    "02 - FEVEREIRO",
    "03 - MARÇO",
    "04 - ABRIL",
    "05 - MAIO",
    "06 - JUNHO",
    "07 - JULHO",
    "08 - AGOSTO",
    "09 - SETEMBRO",
    "10 - OUTUBRO",
    "11 - NOVEMBRO",
    "12 - DEZEMBRO"
]

data_atual = date(2026, 9, 26)
pasta_mes = MESES[data_atual.month - 1]
data_formatada = data_atual.strftime("%d.%m.%Y")

pasta_manha = (
    PASTA_RAIZ
    / str(data_atual.year)
    / pasta_mes
    / data_formatada
    / "MANHÃ"
)

DATA_INICIAL = "23.09.2026"
DATA_FINAL = "24.09.2026"

EMPRESA = "1010"

# ============================================================
# CONFIGURAÇÃO DO CHROME
# ============================================================

options = webdriver.ChromeOptions()

prefs = {
    "download.default_directory": str(pasta_manha),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}

options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

driver.maximize_window()

driver.get(
    "https://sapprd.hec.icatuseguros.com.br/sap/bc/ui2/flp#Shell-home"
)

driver.execute_script("document.body.style.zoom='60%'")


# ============================================================
# LOGIN SAP
# ============================================================
button_avancar = driver.find_element(
    By.CLASS_NAME,
    "urBtnCnt"
)

time.sleep(2)
button_avancar.click()


# E-MAIL
input_email = wait.until(
    EC.presence_of_element_located(
        (By.ID, "i0116")
    )
)

input_email.send_keys(
    "joferreira@icatuseguros.com.br"
)


# AVANÇAR E-MAIL
button_avancar_email = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "idSIButton9")
    )
)

button_avancar_email.click()


# SENHA
def pedir_senha(mensagem="Digite sua senha:"):
    root = tk.Tk()
    root.withdraw()

    senha = simpledialog.askstring(
        "Login SAP",
        mensagem,
        show="*"
    )

    root.destroy()

    return senha


while True:

    senha = pedir_senha(
        mensagem="Digite sua senha:"
    )

    if not senha:
        print("Nenhuma senha informada.")
        break

    input_senha = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "i0118")
        )
    )

    input_senha.send_keys(senha)

    button_avancar_senha = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "idSIButton9")
        )
    )

    button_avancar_senha.click()

    time.sleep(2)

    confirm_senha = driver.find_elements(
        By.ID,
        "i0118"
    )

    if confirm_senha:

        print("Senha incorreta. Digite novamente.")

        confirm_senha[0].clear()

        continue

    break


# CONTINUAR CONECTADO
button_sim = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "idSIButton9")
    )
)

button_sim.click()


# ============================================================
# ABRIR ZCD011 - EMPRESA 1010
# ============================================================

app = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'a[href*="sap-ui2-tcode=ZCD011"]'
        )
    )
)

app.click()

driver.switch_to.frame(
    "application-Shell-startGUI-iframe"
)

driver.execute_script(
    "document.body.style.zoom='60%'"
)


# ============================================================
# PREENCHER PARÂMETROS
# ============================================================

empresa = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[title="Empresa"]'
        )
    )
)

empresa.clear()
empresa.send_keys(EMPRESA)

print(f"Empresa {EMPRESA} preenchida!")


tipo_documento = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[title="Tipo de documento"]'
        )
    )
)

tipo_documento.clear()
tipo_documento.send_keys("56")

print("Tipo de documento preenchido!")


time.sleep(1.5)


data_inicial = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_BUDAT-LOW"]'
        )
    )
)

data_inicial.clear()
data_inicial.send_keys(DATA_INICIAL)

print("Data inicial preenchida!")
time.sleep(1.5)

data_final = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_BUDAT-HIGH"]'
        )
    )
)

data_final.clear()
data_final.send_keys(DATA_FINAL)

print("Data final preenchida!")
time.sleep(1.5)

vencimento_inicial = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_FAEDN-LOW"]'
        )
    )
)

vencimento_inicial.clear()
vencimento_inicial.send_keys(DATA_INICIAL)

print("Vencimento inicial preenchido!")
time.sleep(1.5)

vencimento_final = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_FAEDN-HIGH"]'
        )
    )
)

vencimento_final.clear()
vencimento_final.send_keys(DATA_FINAL)

print("Vencimento final preenchido!")


# ============================================================
# SELEÇÃO MÚLTIPLA
# ============================================================

botao_selecao_multipla = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "M0:46:::19:78")
    )
)

botao_selecao_multipla.click()

tipos_documento = [
    "62",
    "65",
    "79",
    "63"
]


for i, valor in enumerate(tipos_documento):

    campo_id = (
        f"M1:46:1:2B256:1[{i + 2},2]_c"
    )

    campo = wait.until(
        EC.element_to_be_clickable(
            (By.ID, campo_id)
        )
    )

    campo.click()

    time.sleep(1)

    campo = wait.until(
        EC.element_to_be_clickable(
            (By.ID, campo_id)
        )
    )

    campo.send_keys(valor)

    print(
        f"Tipo {valor} preenchido!"
    )


# TRANSFERIR
botao_transferir = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Transferir (F8)"]'
        )
    )
)

botao_transferir.click()
print("Tipos de documento transferidos!")
# ============================================================
# BLOQUEIO DE COMPENSAÇÃO
# ============================================================

time.sleep(2)
while True:

    try:

        radio = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'span[role="radio"][aria-label="Bloqueio de Compensação"]'
                )
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            radio
        )

        print(
            "Bloqueio de Compensação marcado!"
        )

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou o elemento. "
            "Localizando Bloqueio novamente..."
        )


# ============================================================
# EXECUTAR RELATÓRIO
# ============================================================
while True:

    try:

        botao_executar = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'div[role="button"][title=" (F8)"]'
                )
            )
        )

        botao_executar.click()

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou a tela. "
            "Localizando o botão novamente..."
        )


print("Relatório executado!")

# ============================================================
# EXPORTAR EXCEL - EMPRESA 1010
# ============================================================

botao_visualizacao = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Visualização (Ctrl+Shift+F10)"]'
        )
    )
)

botao_visualizacao.click()
print("Modo de exibição aberto!")

time.sleep(2)
while True:

    try:

        botao_file_local = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'div[role="button"][title="File local... (Ctrl+Shift+F9)"]'
                )
            )
        )

        botao_file_local.click()

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou a barra. "
            "Localizando File local novamente..."
        )


print("Menu de exportação aberto!")
time.sleep(1)

planilha = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'span[role="radio"][aria-label="Planilha eletrônica"]'
        )
    )
)

planilha.click()
print("Planilha eletrônica selecionada!")


botao_avancar_exportacao = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Avançar (Entrada)"]'
        )
    )
)

botao_avancar_exportacao.click()
print("Avançar clicado!")


botao_exportar = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Exportar dados (Shift+F8)"]'
        )
    )
)

botao_exportar.click()
print("Exportar para... clicado!")


botao_ok = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "UpDownDialogChoose")
    )
)

botao_ok.click()
print("OK clicado!")
time.sleep(3)

# ============================================================
# VOLTAR PARA HOME
# ============================================================

driver.switch_to.default_content()

botao_inicio = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "shell-header-logo")
    )
)

botao_inicio.click()
print("Página inicial aberta!")

# ============================================================
# ABRIR ZCD011 NOVAMENTE - EMPRESA 1013
# ============================================================
app = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'a[href*="sap-ui2-tcode=ZCD011"]'
        )
    )
)

app.click()

driver.switch_to.frame(
    "application-Shell-startGUI-iframe"
)

driver.execute_script(
    "document.body.style.zoom='60%'"
)

print("ZCD011 aberto novamente!")

# ============================================================
# PREENCHER PARÂMETROS - EMPRESA 1013
# ============================================================

empresa = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[title="Empresa"]'
        )
    )
)

empresa.clear()
empresa.send_keys("1013")

print("Empresa 1013 preenchida!")
time.sleep(1.5)

tipo_documento = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[title="Tipo de documento"]'
        )
    )
)

tipo_documento.clear()
tipo_documento.send_keys("56")

print("Tipo de documento preenchido!")
time.sleep(1.5)


data_inicial = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_BUDAT-LOW"]'
        )
    )
)

data_inicial.clear()
data_inicial.send_keys(DATA_INICIAL)

print("Data inicial preenchida!")
time.sleep(1.5)

data_final = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_BUDAT-HIGH"]'
        )
    )
)

data_final.clear()
data_final.send_keys(DATA_FINAL)
print("Data final preenchida!")
time.sleep(1.5)

vencimento_inicial = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_FAEDN-LOW"]'
        )
    )
)

vencimento_inicial.clear()
vencimento_inicial.send_keys(DATA_INICIAL)

print("Vencimento inicial preenchido!")
time.sleep(1.5)


vencimento_final = wait.until(
    EC.presence_of_element_located(
        (
            By.CSS_SELECTOR,
            'input[lsdata*="S_FAEDN-HIGH"]'
        )
    )
)

vencimento_final.clear()
vencimento_final.send_keys(DATA_FINAL)

print("Vencimento final preenchido!")

# ============================================================
# SELEÇÃO MÚLTIPLA
# ============================================================

while True:

    try:

        botao_selecao_multipla = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "M0:46:::19:78")
            )
        )

        botao_selecao_multipla.click()

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou a tela. "
            "Localizando seleção múltipla novamente..."
        )


tipos_documento = [
    "62",
    "65",
    "63"
]


for i, valor in enumerate(tipos_documento):

    campo_id = (
        f"M1:46:1:2B256:1[{i + 2},2]_c"
    )

    campo = wait.until(
        EC.element_to_be_clickable(
            (By.ID, campo_id)
        )
    )

    campo.click()
    time.sleep(1)

    campo = wait.until(
        EC.element_to_be_clickable(
            (By.ID, campo_id)
        )
    )

    campo.send_keys(valor)

    print(
        f"Tipo {valor} preenchido!"
    )


# TRANSFERIR

botao_transferir = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Transferir (F8)"]'
        )
    )
)

botao_transferir.click()

print("Tipos de documento transferidos!")

# ============================================================
# BLOQUEIO DE COMPENSAÇÃO
# ============================================================

time.sleep(2)

while True:

    try:

        radio = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'span[role="radio"][aria-label="Bloqueio de Compensação"]'
                )
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            radio
        )

        print(
            "Bloqueio de Compensação marcado!"
        )

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou o elemento. "
            "Localizando Bloqueio novamente..."
        )


# ============================================================
# EXECUTAR RELATÓRIO
# ============================================================

while True:

    try:

        botao_executar = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'div[role="button"][title=" (F8)"]'
                )
            )
        )

        botao_executar.click()

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou a tela. "
            "Localizando o botão novamente..."
        )


print("Relatório 1013 executado!")

# ============================================================
# EXPORTAR EXCEL - EMPRESA 1013
# ============================================================

botao_visualizacao = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Visualização (Ctrl+Shift+F10)"]'
        )
    )
)

botao_visualizacao.click()

print("Modo de exibição aberto!")
time.sleep(2)

while True:

    try:

        botao_file_local = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    'div[role="button"][title="File local... (Ctrl+Shift+F9)"]'
                )
            )
        )

        botao_file_local.click()

        break

    except StaleElementReferenceException:

        print(
            "SAP atualizou a barra. "
            "Localizando File local novamente..."
        )


print("Menu de exportação aberto!")
time.sleep(1)


planilha = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'span[role="radio"][aria-label="Planilha eletrônica"]'
        )
    )
)

planilha.click()
print("Planilha eletrônica selecionada!")

botao_avancar_exportacao = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Avançar (Entrada)"]'
        )
    )
)

botao_avancar_exportacao.click()
print("Avançar clicado!")

botao_exportar = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'div[role="button"][title="Exportar dados (Shift+F8)"]'
        )
    )
)

botao_exportar.click()
print("Exportar para... clicado!")

botao_ok = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "UpDownDialogChoose")
    )
)

botao_ok.click()
print("Excel da empresa 1013 exportado!")

# ============================================================
# FINALIZAR
# ============================================================

input("Pressione ENTER manualmente para fechar o navegador...")
driver.quit()