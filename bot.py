from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as ptg
import xlsxwriter as excel

# Lista de compras para pesquisar no site
lista_compra = ['Enxaguante Bucal', 'Desodorante Masculino', 'Shampoo', 'Condicionador', 'Perfume Masculino', 'Escova de Dente']

# Abre o navegador na página inicial do site
navegador = webdriver.Chrome()
navegador.get('https://www.extrabom.com.br/#')

ptg.sleep(2)

# Procura o X da propaganda para fechá-la
navegador.find_element(By.CLASS_NAME, 'spt-close-lightbox').click()

ptg.sleep(2)

# Procura a caixa de pesquisa e digita o produto e aperta ENTER
navegador.find_element(By.NAME, 'q').send_keys('Enxaguante Bucal')
navegador.find_element(By.NAME, 'q').send_keys(Keys.RETURN)

workbook = excel.Workbook(r'C:\Users\Gleidson\OneDrive - Sociedade Educacional do Espírito Santo - UVV\Desktop\PythonES\Jupyter Notebooks\arquivos excel\extrabom srap\lista_de_compras_extrabom.xlsx')
worksheet = workbook.add_worksheet()

# Escrever cabeçalhos
worksheet.write('A1', 'Nome do Produto')
worksheet.write('B1', 'Valor do Produto')

linha = 1  # Começa na linha 2, já que a linha 1 é para os cabeçalhos
for i in range(1, 11):
    # Captura o texto do nome do produto
    nome_do_produto = navegador.find_element(By.CSS_SELECTOR, f'#conteudo div:nth-child(2) div:nth-child(1) div:nth-child({i}) div div div:nth-child(2) .name-produto').text
    
    # Captura o texto do valor
    valor_do_produto = navegador.find_element(By.CSS_SELECTOR, f'#conteudo div:nth-child(2) div:nth-child(1) div:nth-child({i}) div div div:nth-child(2) .item-por__val').text
    
    # Escreve os dados na planilha
    worksheet.write(linha, 0, nome_do_produto)
    worksheet.write(linha, 1, valor_do_produto)
    
    linha += 1

# Feche o navegador
navegador.quit()

# Salvar o arquivo Excel
workbook.close()
