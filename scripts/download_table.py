import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

download_dir = os.path.abspath(os.path.join(os.getcwd(), "downloads"))
os.makedirs(download_dir, exist_ok=True)

print(f"Download directory: {download_dir}", flush=True)

usuario = os.getenv("username")
senha = os.getenv("password")

if not usuario or not senha:
    raise ValueError("Environment variables 'user' and/or 'password' not set.")

# chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--allow-running-insecure-content")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--unsafely-treat-insecure-origin-as-secure=https://sciweb.com.br/")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 100)

def clicar_elemento(xpath):
    try:
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
        driver.execute_script("arguments[0].click();", elemento)
    except Exception as e:
        print(f"Erro ao clicar em {xpath}: {e}", flush=True)


def esperar_download_concluir(nome_arquivo):
    arquivos_iniciais = set(os.listdir(download_dir))
    inicio = time.time()

    while True:
        arquivos_atuais = set(os.listdir(download_dir))
        novos = arquivos_atuais - arquivos_iniciais

        if novos:
            arquivo = novos.pop()
            origem = os.path.join(download_dir, arquivo)
            destino = os.path.join(download_dir, f"{nome_arquivo}.csv")
            os.rename(origem, destino)
            print(f"File saved as: {destino}", flush=True)
            break

        if time.time() - inicio > 60:
            print("Download timeout!", flush=True)
            break

        time.sleep(1)

hoje = datetime.now()
mes = hoje.month + 1
ano = hoje.year

if mes == 13:
    mes = 1
    ano += 1

competencia = f"{mes:02d}/{ano}"
print(f"Competência: {competencia}", flush=True)

xpaths_filiais = [
    f'//*[@id="nav"]/ul/li[14]/ul/li[{i}]/a'
    for i in list(range(1, 12)) + list(range(13, 19))
]

try:
    print("\nInitiating SCI process\n", flush=True)

    # Abre site
    driver.get("https://sciweb.com.br/")

    # Login
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="usuario"]'))).send_keys(usuario)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="senha"]'))).send_keys(senha)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btLoginPrincipal"]'))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rhnetsocial"]'))).click()

    # Loop Filiais exceto 12
    for filial_xpath in xpaths_filiais:
        try:
            index = filial_xpath.split("[")[-1].split("]")[0]

            clicar_elemento(filial_xpath)
            clicar_elemento('//*[@id="menu999"]')
            clicar_elemento('//*[@id="menu9"]')
            clicar_elemento('//*[@id="menu47"]/span[3]')
            clicar_elemento('//*[@id="menu80"]/span[2]')

            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(5)

            clicar_elemento('//input[@id="1-saida" and @value="CSV"]')

            # COMPETENCIA AUTOMÁTICA
            campo = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@id="competencia"]')))
            campo.clear()
            campo.send_keys("12/2025")

            # SELECT2
            select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#s2id_ordenar .select2-choice")))
            select2.click()

            option = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//div[@class='select2-result-label' and contains(normalize-space(), 'Alfabética + Vencimento')]"
            )))
            option.click()

            time.sleep(3)

            clicar_elemento('//button[@type="submit" and contains(text(), "Emitir")]')

            nome_arquivo = f"FERIAS_FILIAL - {index.zfill(2)}"
            esperar_download_concluir(nome_arquivo)

            print(f"OK Filial {index}", flush=True)

        except Exception as e:
            print(f"Erro filial {filial_xpath}: {e}", flush=True)
            
    print("\nRestarting driver for filial 12...\n", flush=True)

    driver.quit()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 100)

    driver.get("https://sciweb.com.br/")
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="usuario"]'))).send_keys(usuario)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="senha"]'))).send_keys(senha)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btLoginPrincipal"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rhnetsocial"]'))).click()

    try:
        filial12 = '//*[@id="nav"]/ul/li[14]/ul/li[12]/a'
        clicar_elemento(filial12)
        clicar_elemento('//*[@id="menu999"]')
        clicar_elemento('//*[@id="menu9"]')
        clicar_elemento('//*[@id="menu47"]/span[3]')
        clicar_elemento('//*[@id="menu80"]/span[2]')

        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(5)

        clicar_elemento('//input[@id="1-saida" and @value="CSV"]')

        campo = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@id="competencia"]')))
        campo.clear()
        campo.send_keys("12/2025")

        select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#s2id_ordenar .select2-choice")))
        select2.click()

        option = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[@class='select2-result-label' and contains(normalize-space(), 'Alfabética + Vencimento')]"
        )))
        option.click()

        clicar_elemento('//button[@type="submit" and contains(text(), "Emitir")]')

        nome_arquivo = "FERIAS_FILIAL - 12"
        esperar_download_concluir(nome_arquivo)

        print("OK Filial 12", flush=True)

    except Exception as e:
        print(f"Error filial 12: {e}", flush=True)

except Exception as e:
    print(f"General error: {e}", flush=True)

finally:
    time.sleep(5)
    driver.quit()
