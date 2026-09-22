from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import time
import tkinter as tk
from tkinter import simpledialog
from selenium.webdriver.common.keys import Keys

from pathlib import Path

pasta_manha = Path(
    r"X:\Comum-PagamentosDiarios\2026\09 - SETEMBRO\26.09.2026\MANHÃ"
)


# CONFIGURAÇÃO DAS DATAS

DATA_INICIAL = "18.09.2026"
DATA_FINAL = "21.09.2026"


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

driver.get("https://sapprd.hec.icatuseguros.com.br/sap/bc/ui2/flp#Shell-home")

driver.execute_script("document.body.style.zoom='70%'")


# AVANÇAR TELA INICIAL SAP

button_avancar = driver.find_element(By.CLASS_NAME, "urBtnCnt")

time.sleep(2)

button_avancar.click()


# CAMPO E-MAIL DO WINDOWS

input_email = wait.until(
    EC.presence_of_element_located((By.ID, "i0116"))
)

input_email.send_keys("joferreira@icatuseguros.com.br")


# AVANÇAR O E-MAIL

button_avancar_email = wait.until(
    EC.element_to_be_clickable((By.ID, "idSIButton9"))
)

button_avancar_email.click()


# CAMPO SENHA DO WINDOWS

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
    senha = pedir_senha(mensagem="Digite sua senha:")

    if not senha:
        print("Nenhuma senha informada.")
        break

    input_senha = wait.until(
        EC.element_to_be_clickable((By.ID, "i0118"))
    )

    input_senha.send_keys(senha)

    button_avancar_senha = wait.until(
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
    )

    button_avancar_senha.click()

    time.sleep(2)

    confirm_senha = driver.find_elements(By.ID, "i0118")

    if confirm_senha:
        print("Senha incorreta. Digite novamente.")
        confirm_senha[0].clear()
        continue

    break


# CONFIRMAÇÃO CONTINUAR CONECTADO

button_sim = wait.until(
    EC.element_to_be_clickable((By.ID, "idSIButton9"))
)

button_sim.click()


# ABRIR ZCD011

app = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'a[href*="sap-ui2-tcode=ZCD011"]')
    )
)

app.click()

driver.switch_to.frame("application-Shell-startGUI-iframe")

driver.execute_script("document.body.style.zoom='70%'")


# PREENCHIMENTO DOS PARÂMETROS ZCD011

empresa = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[title="Empresa"]')
    )
)

empresa.clear()

empresa.send_keys("1010")

print("Empresa preenchida!")


tipo_documento = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[title="Tipo de documento"]')
    )
)

tipo_documento.clear()

tipo_documento.send_keys("56")

print("Tipo de documento preenchido!")


data_inicial = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_BUDAT-LOW"]')
    )
)

data_inicial.clear()

data_inicial.send_keys(DATA_INICIAL)

print("Data inicial preenchida!")


data_final = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_BUDAT-HIGH"]')
    )
)

data_final.clear()

data_final.send_keys(DATA_FINAL)

print("Data final preenchida!")


vencimento_inicial = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_FAEDN-LOW"]')
    )
)

vencimento_inicial.clear()

vencimento_inicial.send_keys(DATA_INICIAL)

print("Vencimento inicial preenchido!")


vencimento_final = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_FAEDN-HIGH"]')
    )
)

vencimento_final.clear()

vencimento_final.send_keys(DATA_FINAL)

print("Vencimento final preenchido!")


# BLOQUEIO DE COMPENSAÇÃO

radio = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'span[role="radio"][aria-label="Bloqueio de Compensação"]')
    )
)

driver.execute_script("arguments[0].click();", radio)

print("Bloqueio de Compensação marcado!")


# EXECUTAR RELATÓRIO

while True:
    try:
        botao_executar = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'div[role="button"][title=" (F8)"]')
            )
        )

        botao_executar.click()

        break

    except StaleElementReferenceException:
        print("SAP atualizou a tela. Localizando o botão novamente...")


print("Relatório executado!")


# MUDAR PARA VISUALIZAÇÃO

botao_visualizacao = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'div[role="button"][title="Visualização (Ctrl+Shift+F10)"]')
    )
)

botao_visualizacao.click()

print("Visualização alterada!")


# ABRIR OPÇÕES DE EXPORTAÇÃO

while True:
    try:
        botao_file_local = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'div[role="button"][title="File local... (Ctrl+Shift+F9)"]')
            )
        )

        botao_file_local.click()

        break

    except StaleElementReferenceException:
        print("SAP atualizou a barra. Localizando File local novamente...")


print("Menu de exportação aberto!")

time.sleep(1)


# SELECIONAR PLANILHA ELETRÔNICA
planilha = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'span[role="radio"][aria-label="Planilha eletrônica"]')
    )
)

planilha.click()

print("Planilha eletrônica selecionada!")


# CLICAR EM AVANÇAR
botao_avancar_exportacao = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'div[role="button"][title="Avançar (Entrada)"]')
    )
)

botao_avancar_exportacao.click()

print("Avançar clicado!")


# CLICAR EM "EXPORTAR PARA..."
botao_exportar = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'div[role="button"][title="Exportar dados (Shift+F8)"]')
    )
)

botao_exportar.click()

print("Exportar para... clicado!")


# CLICAR EM OK
botao_ok = wait.until(
    EC.element_to_be_clickable(
        (By.ID, "UpDownDialogChoose")
    )
)

botao_ok.click()

print("OK clicado!")


input("Pressione ENTER manualmente para fechar o navegador...")