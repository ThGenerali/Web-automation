import openpyxl
import pandas as pd
from PySimpleGUI import PySimpleGUI as sg
import zipfile
import os
import getpass
from selenium.common.exceptions import NoSuchElementException
import datetime
from botcity.web import WebBot, Browser
from botcity.web import By


##servico = Service(ChromeDriverManager().install())
#navegador = webdriver.Chrome(service=servico)
nome_usuario = getpass.getuser()
workbook = openpyxl.Workbook()
bot = WebBot()
lista_questionario = []
pasta_zip = []
diretorio_questionario = [] 
status_survey_questionario = [] 
empresa = ''
aplicativo = True
SITUACAO = 'sucesso'
questionario = ''
janela = True
hoje = str(datetime.datetime.now().strftime('%d.%m.%Y'))
diretorio_relatorio = f''
bot.driver_path = r"C:\\Survey_Monkey\\_internal\\features\\chrome_driver\\chromedriver-win64\\chromedriver.exe"
bot.browser = Browser.CHROME
diretorio_controle = f""
diretorio_controle = f""
diretorio_base_bi = f""
diretorio_base_bi = f""





Survey_link = ''
Conta_Survey_email = ''
Conta_Survey_senha = ''

#Passo 8 - Cria relatório de questionário correspondente a empresa.
def relatorio_questionario(empresa):
    global hoje, status_survey_questionario, diretorio_questionario, diretorio_relatorio, lista_questionario

    # Criar um DataFrame de exemplo
    data = {'Questionários': lista_questionario,
            'Survey': status_survey_questionario,
            'Diretório': diretorio_questionario}
    print(len(lista_questionario))
    print(len(status_survey_questionario))
    print(len(diretorio_questionario))
    df = pd.DataFrame(data)

    
    diretorio_relatorio = f"_internal\\features\\planilha\\relatorio\\{empresa}\\{empresa +'-'+ hoje}.xlsx"

    # Salvar o DataFrame como um arquivo Excel
    df.to_excel(diretorio_relatorio, index=False)

    print(f'O arquivo Excel foi criado em: {diretorio_relatorio}')

#Passo 7 - Chamar a função que lerá a planilha base bi da empresa, colocará os dados dos questionários na planilha base bi e salvará a planilha Bi. 
def ler_planilha_base_bi(empresa):
    global SITUACAO, questionario, diretorio_base_bi, diretorio_base_bi, diretorio_questionario
    
    while True:
        for i in diretorio_questionario:
            try:
                if empresa == '':
                    # Caminho do arquivo Excel existente
                    diretorio_base_bi = diretorio_base_bi
                    # Especificar a planilha na qual você deseja adicionar o conteudo_total
                    nome_planilha = "Planilha1"  # Substitua pelo nome real da sua planilha
                    
                    
                if empresa == '':
                    # Caminho do arquivo Excel existente
                    diretorio_base_bi = diretorio_base_bi
                    # Especificar a planilha na qual você deseja adicionar o conteudo_total
                    nome_planilha = "Base"  # Substitua pelo nome real da sua planilha
                
                # Carregue os arquivos Excel
                df_base_bi = pd.read_excel(diretorio_base_bi, sheet_name=nome_planilha)
                df_questionario = pd.read_excel(i)
                    
                # Adicionar as linhas do df_questionario ao final do df_base_bi
                df_base_bi = pd.concat([df_base_bi, df_questionario.iloc[1:]], ignore_index=True)

                # Salvar o DataFrame atualizado no arquivo1
                df_base_bi.to_excel(diretorio_base_bi, sheet_name=nome_planilha, index=False)
                    
            except FileNotFoundError:
                print(f"Arquivo não encontrado: {i}")
            except Exception as e:
                print(f"Erro ao ler o arquivo {i}: {e}")
        break        

    relatorio_questionario(empresa)
    final_processo_automacao(SITUACAO, empresa)
     
#Passo 6 - Descompacta o zip do questionário para a pasta da empresa correspondente.
def descompacta_zip(empresa):
    global pasta_zip, nome_usuario, download_folder_path, pasta_zip
    # Diretório de destino para a descompactação
    diretorio_destino = fr"C:\Survey_Monkey\_internal\features\planilha\questionario\{empresa}"

    while True:
        for zip in pasta_zip:
            caminho_zip = zip
            # Criar o diretório de destino, se não existir
            if not os.path.exists(diretorio_destino):
                os.makedirs(diretorio_destino)

            # Descompactar o arquivo ZIP
            with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                zip_ref.extractall(diretorio_destino)

            print(f'A pasta ZIP {zip} foi descompactada em: {diretorio_destino}')
        break
    
    
#Passo 5 – Segue o looping do passo anterior e baixa pasta zip e registra o nome da pasta zip e anota o diretorio do questionario.


    
def baixa_questionario(questionario, empresa):
    global status_survey_questionario, nome_usuario
    diretorio_destino = f"_internal\\features\\planilhas\\questionario\\{empresa}"
    #utiliza o try para verificar se o questionário está fechado, caso não o fehca e faz o downloade e caso contrário só prossegue com o download
    try:
        status_button = bot.find_element("//div[@class='wds-type--section-title']//a[contains(@class, 'sm-status-card-survey-status__overall-status') and contains(@class, 'sm-status-card-survey-status__overall-status--open') and text()='ABERTO']", By.XPATH) 
        status_button.click()
        status_survey_questionario.append('Questionário aberto')
        #Abre o layout para fechar o questionário
        close_button = bot.find_element("//td[contains(@class, 'open requires-auth')]//a[contains(@class, 'sm-badge') and contains(text(), 'ABERTO')]", By.XPATH)
        close_button.click()
        #fecha o questionário
        close_button = bot.find_element("//div[contains(@class, 'dialog-btn-bar clearfix')]//a[contains(@class, 'btn red no-border close-collector') and text()='Fechar coletor agora']", By.XPATH)
        close_button.click()
        #vai para a página de download
        start_download_process_button = bot.find_element("//ul[@class='global-navigation-header-tabs-left']/li[contains(@class, 'progressive')]/a[text()='ANALISAR RESULTADOS']", By.XPATH)
        start_download_process_button.click()
        #processo de download do questionário
        download_button = bot.find_element("//div[@global-share-menu and contains(@class, 'persistent-buttons sm-float-r')]//a[contains(@class, 'wds-button wds-button--icon-right wds-button--sm wds-button--arrow-down') and normalize-space(text())='SALVAR COMO']", By.XPATH)
        download_button.click()
        download_button = bot.find_element("//ul[@class='option-menu' and @view-role='actionMenuView']//li[contains(@class, '')]//a[contains(@class, 'option submenu') and normalize-space(text())='Arquivo de exportação']", By.XPATH)
        download_button.click()
        download_button = bot.find_element("//ul[@class='option-menu export-menu analyze-wds' and @view-role='actionMenuView']//li[contains(@class, '')]//a[contains(@class, 'option') and normalize-space(text())='Dados de todas as respostas']", By.XPATH)
        download_button.click()
        #download_button = bot.find_element("//div[@class='section export-preference' and @view-role='ExportDialogTextInputView']//input[@type='text' and @class='sm-input' and @filename='']", By.XPATH)
        #download_button.send_keys(f'{questionario}.zip')
        download_button = bot.find_element("//div[@class='buttons-wrapper']/div[contains(@view-role, 'ExportDialogButtonView')]/a[normalize-space(text())='EXPORTAR']", By.XPATH)
        download_button.click()
        download_button = bot.find_element("//button[@class='nav-tile-btn smf-icon' and @content-key='exports']", By.XPATH)
        download_button.click()
        download_button = bot.find_element("//div[@class='smf-icon item-icon fadeable' and text()='}']", By.XPATH)
        download_button.click()        
        zip_name = bot.find_element("//span[@class='export-name']", By.XPATH)
        print(zip_name.text)
        #In Company-CT_0055.023-Pesquisa de Satisfação-ago/23
        
        pasta_zip.append(f"C:\\Users\\{nome_usuario}\\Downloads\\{zip_name.text}")  
        diretorio_questionario.append(f"{diretorio_destino}\\{zip_name.text}\\{questionario}.xlsx") 
        print(f' diretorio: {diretorio_questionario}') 
        print(f'Pasta_zip: {pasta_zip}')
        
    except:
        print('Questionário fechado')
        status_survey_questionario.append('Questionário fechado')
def config_navegador(questionario, empresa):
#Passo 4 – Abrir o site e ir para o Survey, cria um looping que buscará o questionário e verificará se ele existe (caso exista ou não, será adicionado em uma planilha de relatório de questionário). Se não existir, ele pula para o próximo, mas se existir, ele chama o passo 5 até o looping terminar.
    global navegador, Conta_automacao_link, Conta_automacao_email, Conta_automacao_senha, status_survey_questionario, Survey_link, bot, download_folder_path, navegador, pasta_zip, diretorio_questionario
    
    bot.browse('https://pt.surveymonkey.com/surveys/?ut_source=header')
    login = bot.find_element('//*[@id="username"]', By.XPATH)
    login.send_keys(Conta_Survey_email)
    button_click = bot.find_element('//*[@id="react-app"]/main/div/div/div/section/div/form/div/button', By.XPATH)
    button_click.click()
    senha = bot.find_element('//*[@id="password"]', By.XPATH)
    senha.send_keys(Conta_Survey_senha)
    button_click = bot.find_element('//*[@id="react-app"]/main/div/div/div/section/div/form/div/button', By.XPATH)
    button_click.click()
    sg.popup('''
    Insira manualmente o Código de autenticação.
    Após terminar o processo de autenticação clique em "ok".
            ''')
    while True:
        for i in questionario:
            try:
                print(i)
                pesquisa = bot.find_element('//*[@id="module-active-surveys-new"]/div[1]/span[2]/div/input', By.XPATH)
                pesquisa.send_keys(i) 
                #Pesquisa o questionário
                search_button = bot.find_element("//span[@class='smf-icon start-search']", By.XPATH)
                search_button.click()
                # Se o elemento não for encontrado, adiciona na lista que não foi encontrado e passa para o próximo
                validacao = bot.find_element('//*[@id="module-active-surveys-new"]/div[2]/div[2]/h2[1]', By.XPATH)
                print(validacao.text)
                try:
                    button = bot.find_element('//*[@id="module-active-surveys-new"]/div[1]/span[2]/div/span[2]', By.XPATH)
                    button.click()
                    status_survey_questionario.append('Questionário não encontrado no Survey')
                except:
                    print(f'O questionário {i} não existe')
                    status_survey_questionario.append(f'Questionário:{i}')
                diretorio_questionario.append('Não possui Diretório')
            except:
                try:
                    #Clica no qustionário se houver
                    questionario = bot.find_element(f'//p[@class="notranslate"]/a[contains(@title, "{i}")]', By.XPATH)
                    questionario.click()
                    print('Questionário encontrado')
                    baixa_questionario(i, empresa)
                    bot.browse('https://pt.surveymonkey.com/surveys/?ut_source=header')      
                except:
                    print(f'O questionário {i} não existe')
                    status_survey_questionario.append(f'Questionário:{i}')
                    diretorio_questionario.append('Não possui Diretório') 
        break

            
#        
#Passo 3 – Abrir e ler a planilha controle, pegando a coluna de nome do questionário e do status do questionário (cada coluna atribuída para cada variável - 2). Caso a coluna de nome do questionário for nula, retornará o fim do processo, mostrando que não é capaz de continuar por esse motivo
def lista_titulo_questionarios(empresa):
    global  status_questionario, diretorio_controle, diretorio_controle, lista_questionario

    try:
        if empresa == '':

            # Nome da planilha (sheet) que você deseja selecionar
            nome_planilha = 'Planilha1'
            
            #carregar arquivo excel com pandas
            df = pd.read_excel(diretorio_controle, sheet_name=nome_planilha)
            
            # Selecionar apenas as colunas desejadas ('Pesquisa enviada a Participantes')
            lista_questionario = df['Título'].tolist()

        elif empresa == '':
        # Nome da planilha (sheet) que você deseja selecionar
            nome_planilha = '2024'

            # Carregar o arquivo Excel em um DataFrame do Pandas
            df = pd.read_excel(diretorio_controle, sheet_name=nome_planilha, skiprows=1)

            # Agora você pode acessar a coluna 'Descrição' diretamente
            lista_questionario = df['Descrição']
    except:
        final_processo_automacao('vazia', empresa)
            
    print(len(lista_questionario))

            


#Passo 2 – Confirmar se a planilha de controle está atualizada  
def confirma_atualizacao_arquivo_controle(empresa):
    global aplicativo, janela
    sg.popup(f'''
            A planilha controle {empresa} está atualizada?
            ''')

#Passo extra - mostrar processo de automação ao usuário
def sobre_automacao():
    sg.popup('''
            PASSO 1 - Escolher a empresa para automação.
            PASSO 2 - Confirmar se a planilha de controle está atualizada.
            PASSO 3 - Abrir e ler a planilha controle, pegando a coluna de nome do questionário e do status do questionário (cada coluna atribuída para cada variável - 2). Caso a coluna de nome do questionário for nula, retornará o fim do processo, mostrando que não é capaz de continuar por esse motivo
            PASSO 4 - Abrir o site e ir para o Survey, cria um looping que buscará o questionário e verificará se ele existe (caso exista ou não, será adicionado em uma planilha de relatório de questionário). Se não existir, ele pula para o próximo, mas se existir, ele chama o passo 5 até o looping terminar.
            PASSO 5 - Segue o looping do passo anterior e baixa pasta zip e registra o nome da pasta zip e anota o diretorio do questionario.
            PASSO 6 - Descompacta o zip do questionário para a pasta da empresa correspondente.
            PASSO 7 - Chamar a função que lerá a planilha base bi da empresa, colocará os dados dos questionários na planilha base bi e salvará a planilha Bi. 
            PASSO 8 - Cria relatório de questionário correspondente a empresa.
            PASSO 9 - Atualiza arquivo excel controle.
            PASSO 10 - Abrir um layout notificando que a automação foi realizada, a planilha base bi foi atualizada, informando sobre relatório de questionário da empresa correspondente e pedindo para subir essa planilha para o SharePoint para atualizar online e não ter risco de perder. 
            PASSO 11 - Perguntar se irá fazer outro processo de automação. Caso responda sim, voltar para a tela inicial; caso não, se despedir e fechar a automação.
            ''')
            


#Passo 10 – Abrir um layout notificando que a automação foi realizada, a planilha base bi foi atualizada, informando sobre relatório de questionário da empresa correspondente e pedindo para subir essa planilha para o SharePoint para atualizar online e não ter risco de perder. 
def final_processo_automacao(situacao_coluna_pesquisa, empresa):
    global janela
    if situacao_coluna_pesquisa == 'vazia':
        sg.popup(f'''
                O arquivo excel controle da empresa {empresa} não consta nenhum questionário.
                Por favor, verifique a planilha controle e tente novamente.''')
        janela = False
    else:
        sg.popup(f'''
Processo de automação realizado com sucesso!

O arquivo excel controle e da base BI  da empresa {empresa} foram atualizados.

Foi criado um relatório dos questionários mostrando o nome, caminho e a situação no Survey. O relatório é nomeado com o nome da empresa e o dia em que você realizou a automação.
                
ATENÇÃO: Recomendo verificar se a pasta está atualizada no SharePoint.
                ''')

#Passo 11 – Perguntar se irá fazer outro processo de automação. Caso responda sim, voltar para a tela inicial; caso não, se despedir e fechar a automação.
def pergunta_reuso_programa():
    global janela
    layout = [
        [sg.Column([[sg.Image("C:\\Survey_Monkey\\_internal\\features\\logo\\logo-assinatura.png")]], justification='center')],
        [sg.Column([[sg.Text('A automação acabou de finalizar!', font=('Helvetica', 12, 'bold'))]], justification='center')],
        [sg.Column([[sg.Text('Você deseja utilizar a automação novamente?', font=('Helvetica', 10, 'bold'))]], justification='center')],
        [sg.Column([[sg.Button('Continuar'), sg.Exit('Sair')]])]
    ]

    window = sg.Window("Automação Survey", layout)

    aplicativo = True
    
    while aplicativo == True:
        event, values = window.read()    
        if event in (sg.WINDOW_CLOSED, 'Sair'):
            window.close()
            aplicativo = False
            janela = False
        if event == 'Continuar':
            window.close()
            Tela()
            

#chama as funções para rodar o programa
def start(empresa, empresa):
    global empresa, janela
    
    while janela != False:
        if empresa is True:
            empresa = ''
            confirma_atualizacao_arquivo_controle(empresa)
            lista_titulo_questionarios(empresa)
            config_navegador(lista_questionario, empresa)
            descompacta_zip(empresa)
            ler_planilha_base_bi(empresa)
            pergunta_reuso_programa()
            
            
        elif empresa is True:
            empresa = ''
            confirma_atualizacao_arquivo_controle(empresa)
            lista_titulo_questionarios(empresa)
            config_navegador(lista_questionario, empresa)
            descompacta_zip(empresa)
            ler_planilha_base_bi(empresa)
            pergunta_reuso_programa()
            
#Passo 1 – Escolher a empresa para automação. 
def Tela():
    global aplicativo
    sg.theme('Reddit')
    layout = [
        [sg.Column([[sg.Image("C:\Survey_Monkey\_internal\\features\logo\logo-assinatura.png")]], justification='center')],
        [sg.Column([[sg.Text('Olá! Bem vindo a automação do Survey!', font=('Helvetica', 12, 'bold'))]], justification='center')],
        [sg.Column([[sg.Text('Antes de iniciar certifique-se de que atualizou a planilha de controle na pasta.', font=('Helvetica', 10, 'bold'))]], justification='center')],
        [sg.Column([[sg.Text('Tem dúvido sobre o processo de automação? Clique em "Sobre" para saber passo o passo.'), sg.Button('Sobre')]], justification='center')],
        [sg.Column([[sg.Text('Para qual empresa você deseja fazer o processo?', font=('Helvetica', 10, 'bold'))]], justification='center')], 
        [sg.Column([[sg.Radio('Jornada para o Futuro', key='', group_id='empresa'), sg.Radio('', key='', group_id='empresa')]], justification='center')],
        [sg.Column([[sg.Button('Executar'), sg.Exit('Sair')]])]
    ]

    window = sg.Window("Automação Survey", layout)

    while aplicativo == True:
        event, values = window.read()    
        if event in (sg.WINDOW_CLOSED, 'Sair'):
            aplicativo = False
        
        elif event == 'Executar':
             = values['']
            CD = values['']
            
            if values['']:
                window.close()
                start(, CD)
                
            
            if values['']:
                window.close()
                start(,CD)
                
            
        elif event == 'Sobre':
            sobre_automacao()
            
Tela()