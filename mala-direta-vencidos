#1Bibliotecas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import time

#2. Abre navegador
navegador = webdriver.Chrome()

time.sleep(5)

#3. Abre webmail
navegador.get('https://webmail-seguro.com.br')


#4. verificar se esta logado, caso não esteja logar

seu_usuario = 'meuusuario@dominio' # Trocar pelo usuário
sua_senha =  'password' #Trocar pela senha

if 'Locamail :: Bem-vindo ao Locamail' in navegador.title:
    navegador.find_element('xpath','//*[@id="rcmloginuser"]').send_keys(seu_usuario)
    navegador.find_element('xpath','//*[@id="rcmloginpwd"]').send_keys(sua_senha)
    navegador.find_element('xpath','//*[@id="submitloginform"]').click()
    
elif 'Locamail :: Caixa de entrada' in navegador.title:
    print('caixa de entrada')
else:
    print(f'Você está em {navegador.title}')
    

#4.1 Fechar popup de e-mail seguro
elemento = navegador.find_element('xpath','//*[@id="close-popover"]')

if 'Fechar' in elemento.text:    
    elemento.click()

mala_direta = pd.read_excel('mala_direta_prazo_teste.xlsx')

for index, row in mala_direta.iterrows():
    
    registro = row['Registro']
    empresa = row['Empresa']
    numero_auto = row['Número']
    email  = row['E-mail']
    prazo = 'dd/mm/aaaa' #Substituir pela nova data de prazo
    
    email_corpo = f'''
Sr(a). responsável pela empresa,
{empresa} - {registro}

Lembramos a existência do Auto de Infração {numero_auto}, para qual até a presente data não houve comprovação da regularização da empresa em questão.
Sinalizamos que o prazo para apresentação da comprovação de regularização ou da defesa EXPIROU. 
Contudo, para evitar a adioção de medidas punitivas solicitamos que nos encaminhe comprovação de regularização do auto de infração acima até a data de {prazo}.
''' 
    #Clica em escrever 
    navegador.find_element('xpath', '//*[@id="rcmbtn110"]').click()
    time.sleep(4)
    #6. Na celula de para escrever email do destinatario
    navegador.find_element('xpath','//*[@id="_to"]').send_keys(email)
    time.sleep(1)
    #7. Na celual assunto escrever o assuto: " Aviso de vencimento de prazo
    email_assunto = 'Solicita comprovação de regularização'
    navegador.find_element('xpath','//*[@id="compose-subject"]').send_keys(email_assunto)
    time.sleep(1)

    #8. No corpo do email entrar o texto predefinido

    iframe = navegador.find_element(By.ID, 'composebody_ifr')
    # Muda o contexto para o iframe
    navegador.switch_to.frame(iframe)

    # Localiza o corpo do TinyMCE pelo nome da classe
    tinymce_body = navegador.find_element(By.CLASS_NAME, 'mce-content-body')
    time.sleep(2)
    
    # Limpe o conteúdo atual, se necessário
    tinymce_body.clear()
    # Escreva no TinyMCE
    #tinymce_body.send_keys('Seu texto ou conteúdo aqui.')
    tinymce_body.send_keys(email_corpo)

    #Volta para contexto padrão fora do iframe
    navegador.switch_to.default_content()

    #Envia o e-mail
    navegador.find_element('xpath', '//*[@id="rcmbtn108"]/span').click()
    
    time.sleep(7)
