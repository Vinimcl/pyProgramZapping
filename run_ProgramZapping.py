import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui

webdriver_path = 'C:\\chromedriver-win32\\chromedriver-win32\\chromedriver.exe'
screenshot_folder = 'C:\\Imagens Domicilio\\'
cert_path = 'C:\\Users\\Zapping010\\Certificados\\HERMES E CPF.pfx'
cert_password = '5012'

urls = [
    'https://sso.cloud.pje.jus.br/auth/realms/pje/protocol/openid-connect/auth?client_id=domicilio-eletronico-frontend&redirect_uri=https%3A%2F%2Fdomicilio-eletronico.pdpj.jus.br%2Fselecionar-perfil&state=c0d5f031-b930-493d-8b25-6091cbda6860&response_mode=fragment&response_type=code&scope=openid&nonce=94f73d41-fb48-4e23-83ca-4c89c385dcf7'
]

options = webdriver.ChromeOptions()
options.add_argument(f'--ssl-key-store-type=pkcs12')
options.add_argument(f'--ssl-key-store={cert_path}')
options.add_argument(f'--ssl-key-store-password={cert_password}')

service = Service(webdriver_path)
driver = webdriver.Chrome(service=service, options=options)

try:
    for idx, url in enumerate(urls):
        print(f"Acessando URL: {url}")
        driver.get(url)

        wait = WebDriverWait(driver, 20)

        try:
            print("Aguardando carregamento da página inicial")
            time.sleep(10)

            print("Procurando botão do certificado digital")
            cert_button = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="kc-form-login"]/div/div/div/div[7]/a/span[2]')))
            driver.execute_script("arguments[0].scrollIntoView();", cert_button)

            print("Esperando botão ser clicável")
            cert_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="kc-form-login"]/div/div/div/div[7]/a/span[2]')))
            cert_button.click()
        except Exception as e:
            print(f'Erro ao clicar no botão do certificado digital: {e}')
            continue

        time.sleep(5)

        print("Interagindo com o pop-up do Assinador")
        pyautogui.moveTo(100, 200)
        pyautogui.click()

        time.sleep(5)

        time.sleep(10)

        print("Verificando presença do campo de seleção de empresa...")
        empresa_select_present = EC.presence_of_element_located((By.CLASS_NAME, 'mat-select-value'))
        empresa_select = wait.until(empresa_select_present)
        print("Campo de seleção de empresa encontrado.")

        ActionChains(driver).move_to_element(empresa_select).click().perform()
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'mat-option-text')))

        # Captura todos os elementos de opções de empresa
        empresa_options = driver.find_elements(By.CLASS_NAME, 'mat-option-text')
        num_empresas = len(empresa_options)  # Captura o número total de empresas

        for i in range(num_empresas):
            try:
                print("Selecionando empresa atual")
                ActionChains(driver).move_to_element(empresa_select).click().perform()

                # Use tecla para baixo e Enter para selecionar a próxima empresa
                if i == 0:
                    pyautogui.press('enter')
                else:
                    for _ in range(i+1):
                        pyautogui.press('down')
                    pyautogui.press('enter')
                empresa_nome = empresa_options[i].text  # Captura o nome da empresa selecionada

                print("Usando Tab e Enter para navegar")
                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.press('enter')

                print("Navegando para 'Comunicação Processual'")
                com_proc_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/uikit-layout/mat-sidenav-container/mat-sidenav-content/div[2]/app-home/div/div[3]/mat-card[1]/mat-card-header/div/mat-card-title')))
                com_proc_link.click()

                print("Esperando página carregar")
                time.sleep(5)

                print("Selecionando data")
                start_date_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="start"]')))
                start_date_input.click()
                start_date_input.send_keys('\ue003' * 10)  # Pressiona 'del' 10 vezes
                yesterday_date = (datetime.now() - timedelta(1)).strftime("%d/%m/%Y")
                start_date_input.send_keys(yesterday_date)

                print("Passando para o campo de data final")
                pyautogui.press('tab')  # Passa para o campo de data final
                time.sleep(1)
                pyautogui.press('delete')  # Apaga o conteúdo do campo
                end_date_input = driver.switch_to.active_element  # Garante que estamos no campo de data final
                today_date = datetime.now().strftime("%d/%m/%Y")
                end_date_input.send_keys(today_date)

                print("Clicando no botão 'Buscar'")
                buscar_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/uikit-layout/mat-sidenav-container/mat-sidenav-content/div[2]/app-comunicacoes-layout/app-consulta/div/div[2]/app-filtro/mat-card/mat-card-content/div/div[2]/div[4]/button/span[1]/h5')))
                buscar_button.click()

                print("Esperando resultados carregarem")
                time.sleep(10)

                print("Capturando screenshot")
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                empresa_nome_sanitized = empresa_nome.replace(" ", "_").replace("/", "-")  # Sanitiza o nome da empresa para usar no nome do arquivo
                screenshot_path = os.path.join(screenshot_folder, f'screenshot_{empresa_nome_sanitized}_{timestamp}.png')
                driver.save_screenshot(screenshot_path)
                print(f'Screenshot da empresa salvo em: {screenshot_path}')

                print("Voltando para a página inicial")
                driver.get(url)
                time.sleep(10)

                # Reatualiza as opções de empresa
                empresa_select = wait.until(empresa_select_present)
                ActionChains(driver).move_to_element(empresa_select).click().perform()
                WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'mat-option-text')))
                empresa_options = driver.find_elements(By.CLASS_NAME, 'mat-option-text')

            except Exception as click_err:
                print(f"Erro ao clicar na opção de empresa: {click_err}")

except Exception as e:
    print(f'Erro ao tentar encontrar ou interagir com o campo de seleção de empresa: {e}')

finally:
    driver.quit()
