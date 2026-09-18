from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
from tkinter import simpledialog

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

driver.get("https://sapprd.hec.icatuseguros.com.br/sap/bc/ui2/flp#Shell-home")

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

input("Pressione ENTER para fechar o navegador...")


