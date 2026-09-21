from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import time
import tkinter as tk
from tkinter import simpledialog

# CONFIGURAÇÃO DAS DATAS
DATA_INICIAL = "18.09.2026"
DATA_FINAL = "21.09.2026"

#CONFIGURAÇÕES DO CHROME
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

driver.get("https://sapprd.hec.icatuseguros.com.br/sap/bc/ui2/flp#Shell-home")
driver.execute_script("document.body.style.zoom='70%'")

# AVANÇAR TELA INICIAL SAP
button_avancar = driver.find_element(By.CLASS_NAME, "urBtnCnt")
time.sleep(2)
button_avancar.click()

# CAMPO E-MAIL DO WINDOWS
input_email = wait.until(
    EC.presence_of_element_located((By.ID, "i0116")))
input_email.send_keys("joferreira@icatuseguros.com.br")

# AVANÇAR O E-MAIL
button_avancar_email = wait.until(
    EC.element_to_be_clickable((By.ID, "idSIButton9")))
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
# PASSA A SENHA PRO CHROME

    input_senha.send_keys(senha)
    button_avancar_senha = wait.until(
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
    )

    button_avancar_senha.click()


    time.sleep(2)
    confirm_senha = driver.find_elements(By.ID, "i0118")

    if confirm_senha:
# SENHA INCORRETA
        print("Senha incorreta. Digite novamente.")

# LIMPA O CAMPO E SOLICITA NOVAMENTE
        confirm_senha[0].clear()

        continue
    break

# CONFIRMAÇÃO CONTINUAR CONECTADO
button_sim = wait.until(
    EC.element_to_be_clickable((By.ID, "idSIButton9")))

button_sim.click()

# ABRIR APLICATIVO DESBLOQUEIO DOCUMENTOS DE PAGAMENTO ZCD011
app = wait.until(
    EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'a[href*="sap-ui2-tcode=ZCD011"]')
    )
)
app.click()

driver.switch_to.frame("application-Shell-startGUI-iframe")
driver.execute_script("document.body.style.zoom='70%'")

# PREENCHIMENTO DOS PARÂMETROS ZCD011
# EMPRESA
empresa = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[title="Empresa"]')
    )
)

empresa.clear()
empresa.send_keys("1010")
print("Empresa preenchida!")

# TIPO DE DOCUMENTO
tipo_documento = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[title="Tipo de documento"]')
    )
)

tipo_documento.clear()
tipo_documento.send_keys("56")

print("Tipo de documento preenchido!")

# DATA LANÇAMENTO INICIAL
data_inicial = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_BUDAT-LOW"]')
    )
)

data_inicial.clear()
data_inicial.send_keys(DATA_INICIAL)

print("Data inicial preenchida!")

# DATA LANÇAMENTO FINAL
data_final = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_BUDAT-HIGH"]')
    )
)

data_final.clear()
data_final.send_keys(DATA_FINAL)

print("Data final preenchida!")

# DATA DE VENCIMENTO INICIAL
vencimento_inicial = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_FAEDN-LOW"]')
    )
)

vencimento_inicial.clear()
vencimento_inicial.send_keys(DATA_INICIAL)
print("Vencimento inicial preenchido!")

# DATA DE VENCIMENTO FINAL
vencimento_final = wait.until(
    EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[lsdata*="S_FAEDN-HIGH"]')
    )
)

vencimento_final.clear()
vencimento_final.send_keys(DATA_FINAL)
print("Vencimento final preenchido!")

# MARCAR "BLOQUEIO DE COMPENSAÇÃO"
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
input("Pressione ENTER para encerrar...")