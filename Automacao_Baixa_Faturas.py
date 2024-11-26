
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import os
import time
import pyautogui  # Importa a biblioteca para captura de tela
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException


# Função para verificar se o Acesso é Legado
def verificar_acesso_legado(browser):
    try:
        # Espera o elemento "Acesso Legado" estar presente
        acesso_legado = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.XPATH, "//option[@un='Acesso Legado']"))
        )
        return True  # Retorna True se Acesso Legado for encontrado
    except TimeoutException:
        return False  # Retorna False caso não encontre "Acesso Legado"


# Função para tirar screenshot da tela inteira
def tirar_screenshot(nome_arquivo):
    screenshot_folder = 'C:\\Users\\F8089999\\OneDrive - TIM\\F8089999\\Screenshots'
    if not os.path.exists(screenshot_folder):
        os.makedirs(screenshot_folder)
    screenshot_path = os.path.join(screenshot_folder, nome_arquivo)
    pyautogui.screenshot(screenshot_path)  # Captura a tela inteira
    print(f"Screenshot salvo: {screenshot_path}")


# Carregar o arquivo Excel no DataFrame
df = pd.read_excel(r'C:\Users\F8089999\OneDrive - TIM\_Consolidado Testes Pós Produção de '
                   r'Ofertas\Baixa de Faturas\Auto\Automações - Quality Assurance\Faturas.xlsx')


# Acessar Siebel
def acessar_siebel(browser):
    try:
        app_nome_element = browser.find_element(By.XPATH,
                                                "//div[@class='app_nome' and contains(text(), 'SIEBEL-POS')]")
        link_element = app_nome_element.find_element(By.XPATH,
                                                    "./ancestor::a[@class='link_aplicacao']")
        link_element.click()
        time.sleep(10)
        print("Acesso ao Siebel concluído.")
    except NoSuchElementException:
        print("Elemento Siebel não encontrado.")


# Configurar as opções do Chrome para desabilitar notificações
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-notifications")

# Launch Google Chrome com as opções configuradas
browser = webdriver.Chrome(options=chrome_options)
browser.maximize_window()

# Acessar o PAC PORTAL
browser.get("https://pacportal.internal.timbrasil.com.br/pac/")
time.sleep(3)

# Mapeamento e login no PAC PORTAL
user = browser.find_element(By.ID, "USER")
password = browser.find_element(By.ID, "PASSWORD")
user.send_keys("F8089999")
password.send_keys("Ag#120922")
browser.find_element(By.ID, "btnSubmit").click()
time.sleep(5)


# Acessando Siebel Pós
acessar_siebel(browser)

desired_title = "SmartID"
for handle in browser.window_handles:
    browser.switch_to.window(handle)
    if browser.title == desired_title:
        print("Switched to tab:", browser.title)
        break

# Realizando 2º login - Selenium
usuario = browser.find_element(By.ID, "username")
senha = browser.find_element(By.ID, "password")
usuario.send_keys("F8089999")
senha.send_keys("Vn#120922")
browser.find_element(By.ID, "signOnButton").click()

# Adicionar um tempo para aguardar a página carregar completamente
time.sleep(10)  # Aumentar o tempo de espera se necessário

# Iterar sobre cada linha no DataFrame
for index, row in df.iterrows():
    access_code = int(row['Número'])  # Coluna "Número"
    automation_text = str(row['TEXTO AUTOMAÇÃO'])  # Coluna "TEXTO AUTOMAÇÃO"
   
    # Exibir as informações lidas
    print(f"Processed linha {index}: Código de acesso {access_code}, Texto de automação: {automation_text}")


    # CLICAR EM ACESSOS
    desired_title = "Siebel CRC Pos"
    for handle in browser.window_handles:
        browser.switch_to.window(handle)
        if browser.title == desired_title:
            print("Switched to tab:", browser.title)
            break
    try:
        wait = WebDriverWait(browser, 15)  # Aumentar o tempo de espera se necessário
        element_acessos = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Acessos']/parent::li"))
        )
        element_acessos.click()
        element_acessos.click()
    except NoSuchElementException:
        print("Elemento 'Acessos' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    time.sleep(10)


    # Clicar em Pesquisar Acessos
    try:
        wait = WebDriverWait(browser, 15)  # Aumentar o tempo de espera se necessário
        element_pesquisar = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="s_1_1_108_0_Ctrl"]'))
        )
        element_pesquisar.click()  # Clica uma vez
        element_pesquisar.click()  # Clica novamente, se necessário

    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


    # Localiza o campo e escreve o código de acesso (GSM) obtido no DataFrame
    digitar_acesso = browser.find_element(By.NAME, "s_1_1_44_0")
    digitar_acesso.send_keys(access_code)  # Corrigido para usar a variável access_code

    # Clicar em IR
    try:
        wait = WebDriverWait(browser, 15)
        element_ir = wait.until(
            EC.element_to_be_clickable((By.ID,"s_1_1_97_0_Ctrl"))
        )
        element_ir.click()  # Clica uma vez
        element_ir.click()  # Clica novamente, se necessário

    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        time.sleep(3)

    # Clicar no alerta depois do ir
    try:
        # Aguarda até o alerta aparecer (com timeout de 10 segundos, por exemplo)
        WebDriverWait(browser, 60).until(EC.alert_is_present())
       
        # Alterna para o alerta e o aceita
        alerta = Alert(browser)
        alerta.accept()
        print("Alerta aceito.")
    except (NoAlertPresentException, TimeoutException):
        print("Nenhum alerta presente ou o tempo de espera expirou.")

    # Verificar se "Acesso Legado" está presente
    try:
        acesso_legado = WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.XPATH, "//option[@un='Acesso Legado']"))
        )
        if acesso_legado:
            print(f"Acesso {access_code} é Legado (Pré Pago). Pulando para o próximo.")
            continue
    except TimeoutException:
        print(f"Acesso {access_code} não é Legado. Continuando.")


    # Rola até o final da página para exibir os botões: Ofertas e Histórico de atendimento
    for _ in range(10):
        pyautogui.press('down')


    # Clicar em Ofertas
    try:
        # Aguarda até que o elemento "Ofertas" esteja visível e clicável
        wait = WebDriverWait(browser, 15)  # Aumente o tempo de espera, se necessário
        element_ofertas = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Ofertas']"))
        )
        element_ofertas.click()  # Clica uma vez
        element_ofertas.click()  # Clica novamente, se necessário

    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        tirar_screenshot('home_page_siebel.png')  # Captura após acessar a hp do siebel
        time.sleep(5)

    # Clicar no alerta
    try:
        # Aguarda até o alerta aparecer (com timeout de 10 segundos, por exemplo)
        WebDriverWait(browser, 60).until(EC.alert_is_present())
       
        # Alterna para o alerta e o aceita
        alerta = Alert(browser)
        alerta.accept()
        print("Alerta aceito.")
    except (NoAlertPresentException, TimeoutException):
        print("Nenhum alerta presente ou o tempo de espera expirou.")

    time.sleep(5)
    # Rola até o final da página para exibir os botões: Ofertas e Histórico de atendimento
    for _ in range(10):
        pyautogui.press('down')


    # Clicar em histórico de atendimento
    try:
        wait = WebDriverWait(browser, 30)  # Aumentar o tempo de espera se necessário
        element_historico = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Histórico Atendimento']"))
        )
        element_historico.click()  # Clica uma vez
        element_historico.click()  # Clica novamente, se necessário

    except NoSuchElementException:
        print("Elemento 'Historico' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        time.sleep(5)

    try:
        # Aguarda até o alerta aparecer (com timeout de 10 segundos, por exemplo)
        WebDriverWait(browser, 60).until(EC.alert_is_present())
       
        # Alterna para o alerta e o aceita
        alerta = Alert(browser)
        alerta.accept()
        print("Alerta aceito.")
    except (NoAlertPresentException, TimeoutException):
        print("Nenhum alerta presente ou o tempo de espera expirou.")

    # Rola até o final da página para exibir o botão: Novo
    for _ in range(10):
        pyautogui.press('down')

    # Clicar em novo para abrir tripleta
    try:
        wait = WebDriverWait(browser, 30)  # Aumentar o tempo de espera se necessário
        element_novo = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-display='Novo']"))
        )
        element_novo.click()  # Clica uma vez

    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(4)

    # Preencher tripleta - Em andamento
    try:
        wait = WebDriverWait(browser, 15)  # Aumentar o tempo de espera se necessário
        element_novo = wait.until(
            EC.element_to_be_clickable((By.XPATH, "(//td[@aria-roledescription='Motivo 1'])[1]"))
        )
        element_novo.click()  # Clica uma vez
        element_novo.click()  # Clica uma vez

    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(3)

    # Motivo 1 - Processo Interno
    def selecionar_motivo_1(browser):
        wait = WebDriverWait(browser, 30)

        # Localiza e clica no campo "Motivo 1"
        campo_motivo_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[@aria-roledescription='Motivo 1']")))
        campo_motivo_1.click()

        # Verifica se o campo está habilitado e visível
        if campo_motivo_1.is_enabled() and campo_motivo_1.is_displayed():
            campo_motivo_1.click()  # Clica no campo para focá-lo
            actions = ActionChains(browser)
            actions.click(campo_motivo_1).click(campo_motivo_1).click(campo_motivo_1).perform()  # Simula três cliques
            actions.send_keys(Keys.BACKSPACE).perform()  # Limpa o texto selecionado
            # Escreve "Processo Interno" no campo
            actions.send_keys("Processo Interno").perform()  
            print("Campo atualizado com sucesso.")
        else:
            print("Erro: O campo não está habilitado ou visível.")

    # Exemplo de uso
    selecionar_motivo_1(browser)

    time.sleep(3)

    # Motivo 2 - Conta
    def selecionar_motivo_2(browser):
        wait = WebDriverWait(browser, 30)

        # Localiza e clica no campo "Motivo 2"
        campo_motivo_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[@aria-roledescription='Motivo 2']")))
        campo_motivo_2.click()

        # Verifica se o campo está habilitado e visível
        if campo_motivo_2.is_enabled() and campo_motivo_2.is_displayed():
            campo_motivo_2.click()  # Clica no campo para focá-lo
            actions = ActionChains(browser)
            actions.click(campo_motivo_2).click(campo_motivo_2).click(campo_motivo_2).perform()  # Simula três cliques
            actions.send_keys(Keys.BACKSPACE).perform()  # Limpa o texto selecionado
            # Escreve "Conta" no campo
            actions.send_keys("Conta").perform()  
            print("Campo atualizado com sucesso.")
        else:
            print("Erro: O campo não está habilitado ou visível.")

    # Exemplo de uso
    selecionar_motivo_2(browser)

    time.sleep(3)

    # Motivo 3 - Ajuste Fatura Num Teste
    def selecionar_motivo_3(browser):
        wait = WebDriverWait(browser, 30)

        # Localiza e clica no campo "Motivo 3"
        campo_motivo_3 = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[@aria-roledescription='Motivo 3']")))
        campo_motivo_3.click()

        # Verifica se o campo está habilitado e visível
        if campo_motivo_3.is_enabled() and campo_motivo_3.is_displayed():
            campo_motivo_3.click()  # Clica no campo para focá-lo
            actions = ActionChains(browser)
            actions.click(campo_motivo_3).click(campo_motivo_3).click(campo_motivo_3).perform()  # Simula três cliques
            actions.send_keys(Keys.BACKSPACE).perform()  # Limpa o texto selecionado
            # Escreve "Ajuste Fatura Num Teste" no campo
            actions.send_keys("Ajuste Fatura Num Teste").perform()  
            print("Campo atualizado com sucesso.")
        else:
            print("Erro: O campo não está habilitado ou visível.")

    # Exemplo de uso
    selecionar_motivo_3(browser)

    time.sleep(3)

    # Clicar em prosseguir
    try:
        wait = WebDriverWait(browser, 60)  # Aumentar o tempo de espera se necessário
        element_prosseguir = wait.until(
            browser.find_element(By.XPATH, "//button[@data-display='Prosseguir']").click()
        )
        element_prosseguir.click()  # Clica uma vez
        element_prosseguir.click()  # Clica novamente, se necessário
    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    # Clicar no alerta 01
    try:
        # Aguarda até o alerta aparecer (com timeout de 10 segundos, por exemplo)
        WebDriverWait(browser, 60).until(EC.alert_is_present())

        # Alterna para o alerta e o aceita
        alerta = Alert(browser)
        alerta.accept()
        print("Alerta aceito.")
    except (NoAlertPresentException, TimeoutException):
        print("Nenhum alerta presente ou o tempo de espera expirou.")

    time.sleep(10)

    # Clicar no alerta 02
    try:
        # Aguarda até o alerta aparecer (com timeout de 10 segundos, por exemplo)
        WebDriverWait(browser, 60).until(EC.alert_is_present())

        # Alterna para o alerta e o aceita
        alerta = Alert(browser)
        alerta.accept()
        print("Alerta aceito.")
    except (NoAlertPresentException, TimeoutException):
        print("Nenhum alerta presente ou o tempo de espera expirou.")

    time.sleep(5)

    # Clicar em Detalhe
    try:
        wait = WebDriverWait(browser, 60)  # Aumentar o tempo de espera se necessário
        element_detalhe = wait.until(
            browser.find_element(By.XPATH, "//button[@data-display='Detalhe']").click()
        )
        element_detalhe.click()  # Clica uma vez
        element_detalhe.click()  # Clica novamente, se necessário
    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(5)

    # Clicar em Solicitações de Serviço
    try:
        wait = WebDriverWait(browser, 15)  # Aumentar o tempo de espera se necessário
        element_solicitacao = wait.until(
            browser.find_element(By.XPATH, browser.find_element(By.CSS_SELECTOR, "a[data-tabindex='tabView1']").click())
        )
        element_solicitacao.click()  # Clica uma vez
        element_solicitacao.click()  # Clica novamente, se necessário
    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(5)

    # Clicar em adicionar notas
    try:
        wait = WebDriverWait(browser, 15)  # Aumentar o tempo de espera se necessário
        element_adicionar_notas = wait.until(
            browser.find_element(By.XPATH, "//button[@data-display='Adicionar Notas']").click()
        )
        element_adicionar_notas.click()  # Clica uma vez
        element_adicionar_notas.click()  # Clica novamente, se necessário
    except NoSuchElementException:
        print("Elemento 'Pesquisar' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(5)

    # Escrever nota
    # Iterar sobre cada linha no DataFrame
    # Localiza o campo e escreve o código de acesso (GSM) ou texto de nota
    campo_texto_nota = browser.find_element(By.NAME, "s_1_1_0_0")
    campo_texto_nota.send_keys(automation_text)  # Substitua 'access_code' pelo conteúdo que deseja inserir

    time.sleep(5)

    # Clicar em gravar nota
    # Localiza o botão "Gravar Nota" e clica
    botao_gravar_nota = browser.find_element(By.XPATH, "//button[@data-display='Gravar Nota']")
    botao_gravar_nota.click()

    # Ir até em baixo
    # Rola até o final da página para exibir o botão: Novo
    for _ in range(10):
        pyautogui.press('down')

    time.sleep(5)

    # Escolher arquivo
    # Localiza o campo de upload de arquivo e envia o caminho completo do arquivo
    campo_upload = browser.find_element(By.ID, "s_2_1_5_0s_SweFileName")
    campo_upload.send_keys("C:/Users/F8089999/OneDrive - TIM/_Consolidado Testes Pós Produção de Ofertas/Baixa de Faturas/2024/De Acordo Novembro/11. Re Pós Rollout Baixa Faturas - Linhas de teste - Nov24.msg")

    time.sleep(5)

    # Ir até o inicio da tela
    # Rola até em cima da página para exibir o botão: Encaminhar
    for _ in range(10):
        pyautogui.press('up')

    time.sleep(5)

    # Clicar em encaminhar
    # Localiza o botão "Encaminhar" e clica nele
    botao_encaminhar = browser.find_element(By.XPATH, "//button[@data-display='Encaminhar']")
    botao_encaminhar.click()

    # Clicar no alerta final após clique em encaminhar  
    try:
        # Alterna para o alerta
        alerta = Alert(browser)
        # Clica no botão "OK" (ou equivalente) do alerta
        alerta.accept()
        print("Alerta aceito.")
    except NoAlertPresentException:
        print("Nenhum alerta presente.")

    time.sleep(3)

    nome_print = "protocolo_gsm_" + str(access_code) + ".png"
    print('NOME ARQUIVO:', nome_print)
    tirar_screenshot(nome_print)

    protocolo_gsm = "CHAMADO ABERTO PARA O GSM"  + str(access_code) +""
    print (protocolo_gsm)

#teste