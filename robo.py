from time import sleep
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chrome_path = 'C:/Users/Usuario/Desktop/ChromeDriver/chromedriver-win64/chromedriver.exe'
chrome_service = webdriver.chrome.service.Service(chrome_path)
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {"download.default_directory": "C:/Users/Usuario/Desktop/MepCred"})

try:
    n = webdriver.Chrome(service=chrome_service, options=chrome_options)
except Exception as e:
    print(f"An error occurred: {e}")
    n = webdriver.Chrome(options=chrome_options)

n.get("https://picapau.info/users/login")
sleep(2)
n.find_element("xpath", '//*[@id="id_email"]').send_keys("mepcred@gmail.com")
sleep(2)
n.find_element("xpath", '//*[@id="id_password"]').send_keys("102030")
sleep(2)
n.find_element("xpath", '/html/body/div[1]/section[2]/div/div/div/form/div[2]/input').click()
sleep(2)
n.find_element("xpath", '/html/body/div[2]/div/div/a[1]/div').click()

csv_path = "C:/Users/Usuário/Desktop/testeBase.csv"
df = pd.read_csv(csv_path, encoding='utf-8')

# Criar DataFrame fora do loop
resultados_df = pd.DataFrame()

for index, row in df.iterrows():
    # Preencher campo CPF com o valor da linha atual
    if 'teste' in df.columns:
        cpf_input = n.find_element("xpath", '//*[@id="id_cpf_nb"]')
        cpf_input.clear()
        cpf_value = str(row['teste'])  # Substitua 'teste' pelo nome real da coluna no CSV
        cpf_input.send_keys(cpf_value)

        # Clicar no botão de pesquisa
        n.find_element("xpath", '//*[@id="searchForm"]/div/div/input[2]').click()
        # Aguardar um pouco para os resultados serem processados
        sleep(5)
        # n.find_element("xpath", '//*[@id="search-actions"]/div/form/button').click()
        sleep(10)  # Pode ser ajustado conforme necessário

        elemento_1 = WebDriverWait(n, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="result"]/div[2]/fieldset[3]/p[1]')))
        sleep(5)
        elemento_2 = WebDriverWait(n, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="result"]/div[2]/fieldset[3]/p[3]')))
        sleep(5)
        elemento_3 = WebDriverWait(n, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="result"]/div[2]/fieldset[7]/p[5]')))
        sleep(5)
        elemento_4 = WebDriverWait(n, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="result"]/div[2]/fieldset[7]/p[4]')))
        sleep(5)
        elemento_5 = WebDriverWait(n, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="result"]/div[2]/fieldset[7]/p[6]')))
        sleep(5)

        texto_extraido1 = elemento_1.text
        texto_extraido2 = elemento_2.text
        texto_extraido3 = elemento_3.text
        texto_extraido4 = elemento_4.text
        texto_extraido5 = elemento_5.text


        # Adicionar dados ao DataFrame
        resultados_df = pd.concat([resultados_df, pd.DataFrame({'Informaçãoes extraídas 1': [texto_extraido1],
                                                                'Infomações extraidas 2': [texto_extraido2],
                                                                'Informações extraidas 3': [texto_extraido3],
                                                                'Informações extraidas 4': [texto_extraido4],
                                                                'Informações extraidas 5': [texto_extraido5]})])

# Salvar o DataFrame em um arquivo Excel
excel_path = "C:/Users/Usuário/Desktop/informacoesRobo.xlsx"
resultados_df.to_excel(excel_path, index=False)

# Fechar o navegador
print("Informações extraídas com sucesso!")

try:
    # Tente fechar o navegador
    n.quit()
except WebDriverException as e:
    # Captura a exceção e imprime uma mensagem de erro
    print(f"Erro ao fechar o navegador: {e}")

# Mensagem antes de encerrar o programa
print("Programa encerrado.")

        # n.find_element("xpath", '//*[@id="result"]/div[3]').click()
        # sleep(10)
        # img_url = n.get_attribute("src")
        # response = requests.get(img_url)
        # with open("C:/Users/Usuario/Desktop/MepCred","wb") as f:
        #     f.write(response.content)

# Fechar o navegador