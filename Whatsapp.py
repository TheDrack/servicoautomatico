from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
import time
import urllib
from tkinter import messagebox
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import os

PlanilhaFila = []

def Interface():
    global janela
    global botao_iniciar
    global TempoMaximo
    #Abrir uma interface para iniciar ou parar a leitura da fila
    janela = Tk()
    janela.title("WhatsApp")
    janela.geometry("300x100")

    #botao criar fila
    botao_criar_fila = Button(janela, text="Criar Fila", command=CriarFila)
    botao_criar_fila.place(x=50, y=50)

    botao_iniciar = Button(janela, text="Iniciar Disparo", command=LerFila)
    botao_iniciar.place(x=150, y=50)

    #colocar um campo de tempo maximo inativo no disparo
    label = Label(janela, text="Tempo Maximo Inativo")
    label.place(x=50, y=20)
    TempoMaximo = Entry(janela, width=10)
    TempoMaximo.place(x=180, y=20)

    janela.mainloop()

def CriarFila():
    global PlanilhaFila
    PlanilhaFila = []
    # Interface gráfica para adicionar uma nova execução
    def adicionar_execucao_gui():
        global arquivo_selecionado
        global JanelaFila
        # Cria a JanelaFila
        JanelaFila = tk.Tk()
        JanelaFila.title("Adicionar execução")
        JanelaFila.geometry("300x170")


        # Botão de selecionar arquivo
        botao_selecionar_arquivo = tk.Button(JanelaFila, text="Selecionar arquivo", command=selecionar_arquivo)
        botao_selecionar_arquivo.grid(row=0, column=0, padx=10, pady=10)

        # Botão de adicionar execução
        botao_adicionar_execucao = tk.Button(JanelaFila, text="Adicionar execução", command=adicionar_execucao)
        botao_adicionar_execucao.grid(row=1, column=0, padx=10, pady=10)

        # Inicia a JanelaFila
        JanelaFila.mainloop()


    # Função para selecionar o arquivo
    def selecionar_arquivo():
        global arquivo_selecionado
        global colunas
        global OpcaoNome
        global OpcaoCpf
        global OpcaoTelefone1
        global OpcaoTelefone2
        global OpcaoTelefone3
        global OpcaoValor
        global pastaArquivoSelecionado
        
        # Seleciona o arquivo
        arquivo_selecionado = filedialog.askopenfilename(initialdir='I:\__GIT_jesus__\Listas Disparos', title="Selecione o arquivo", filetypes=(("Excel", "*.xlsx"), ("Excel", "*.xls"), ("Excel", "*.xlsm")))
        #pasta do arquivo selecionado
        pastaArquivoSelecionado = os.path.dirname(arquivo_selecionado)

        #definir as colunas do arquivo selecionado, para quais serão de suas determinadas variaveis, nome, cpf, telefone1, telefone2, telefone3, valor
        colunas = ["Nome", "CPF", "Telefone 1", "Telefone 2", "Telefone 3", "Valor"]

        # Lê o arquivo
        arquivo = pd.read_excel(arquivo_selecionado)
        #cria na JanelaFila, dropdown com as colunas do arquivo selecionado, para selecionar qual coluna será de cada variavel

        LabelNome = tk.Label(JanelaFila, text="Nome")
        LabelNome.grid(row=0, column=1, padx=10, pady=10)

        OpcaoNome = tk.StringVar(JanelaFila)
        OpcaoNome.set("Selecionar")
        coluna_nome = tk.OptionMenu(JanelaFila, OpcaoNome, *arquivo.columns)
        coluna_nome.grid(row=1, column=1, padx=10, pady=10)

        LabelCpf = tk.Label(JanelaFila, text="CPF")
        LabelCpf.grid(row=0, column=2, padx=10, pady=10)

        OpcaoCpf = tk.StringVar(JanelaFila)
        OpcaoCpf.set("Selecionar")
        coluna_cpf = tk.OptionMenu(JanelaFila, OpcaoCpf, *arquivo.columns)
        coluna_cpf.grid(row=1, column=2, padx=10, pady=10)

        LabelTelefone1 = tk.Label(JanelaFila, text="Telefone 1")
        LabelTelefone1.grid(row=0, column=3, padx=10, pady=10)

        OpcaoTelefone1 = tk.StringVar(JanelaFila)
        OpcaoTelefone1.set("Selecionar")
        coluna_telefone1 = tk.OptionMenu(JanelaFila, OpcaoTelefone1, *arquivo.columns)
        coluna_telefone1.grid(row=1, column=3, padx=10, pady=10)

        LabelTelefone2 = tk.Label(JanelaFila, text="Telefone 2")
        LabelTelefone2.grid(row=0, column=4, padx=10, pady=10)

        OpcaoTelefone2 = tk.StringVar(JanelaFila)
        OpcaoTelefone2.set("Selecionar")
        coluna_telefone2 = tk.OptionMenu(JanelaFila, OpcaoTelefone2, *arquivo.columns)
        coluna_telefone2.grid(row=1, column=4, padx=10, pady=10)

        LabelTelefone3 = tk.Label(JanelaFila, text="Telefone 3")
        LabelTelefone3.grid(row=0, column=5, padx=10, pady=10)

        OpcaoTelefone3 = tk.StringVar(JanelaFila)
        OpcaoTelefone3.set("Selecionar")
        coluna_telefone3 = tk.OptionMenu(JanelaFila, OpcaoTelefone3, *arquivo.columns)
        coluna_telefone3.grid(row=1, column=5, padx=10, pady=10)

        LabelValor = tk.Label(JanelaFila, text="Valor")
        LabelValor.grid(row=0, column=6, padx=10, pady=10)

        OpcaoValor = tk.StringVar(JanelaFila)
        OpcaoValor.set("Selecionar")
        coluna_valor = tk.OptionMenu(JanelaFila, OpcaoValor, *arquivo.columns)
        coluna_valor.grid(row=1, column=6, padx=10, pady=10)

        #ajustar tamanho da JanelaFila para caber os dropdowns
        JanelaFila.geometry("1000x170")
        JanelaFila.focus_force()

    # Função para adicionar uma execução
    def adicionar_execucao():
        global arquivo_selecionado
        global CaixaDeMensagem

        colunasSelecionadas = [OpcaoNome.get(), OpcaoCpf.get(), OpcaoTelefone1.get(), OpcaoTelefone2.get(), OpcaoTelefone3.get(), OpcaoValor.get()]


        arquivo = pd.read_excel(arquivo_selecionado)
        caminho = f'{pastaArquivoSelecionado}/mensagem.txt'
        # Adiciona cada linha do arquivo na fila, com base nas colunas selecionadas
        for index, row in arquivo.iterrows():
            PlanilhaFila.append((row[colunasSelecionadas[0]], row[colunasSelecionadas[1]], row[colunasSelecionadas[2]], row[colunasSelecionadas[3]], row[colunasSelecionadas[4]], row[colunasSelecionadas[5]]))
        JanelaFila.destroy()
        MensagemLabel = tk.Label(janela, text="Mensagem:")
        MensagemLabel.place(x=10, y=80)
        CaixaDeMensagem = tk.Text(janela, width=100, height=10)
        CaixaDeMensagem.place(x=10, y=100)
        #se existir um arquivo de mensagem.txt, preenche a caixa de mensagem com o conteudo do arquivo
        if os.path.exists(caminho):
            with open(caminho, 'r') as file:
                CaixaDeMensagem.insert(tk.END, file.read())
        janela.geometry("1000x300")
    adicionar_execucao_gui()



driverWhats = None

def inicializacao():
    global driverWhats
    global options

    options = Options()
    driverWhats = webdriver.Firefox(options=options)

    # options.add_argument('-headless')
    driverWhats.get('https://web.whatsapp.com/')
    driverWhats.maximize_window()
    time.sleep(5)
    #aguardar até conectar o QR cODE
    while True:
        try: driverWhats.find_element(By.XPATH,'//*[@id="side"]')
        except: time.sleep(1)
        else: break

    Interface()

def enviar_mensagem(numero, contato, cpf, valor):
    telefone = str(numero).replace('.0','')
    contato = str(contato).split(' ')[0]
    mensagem = CaixaDeMensagem.get("1.0", tk.END)
    #criar um txt para salvar a mensagem


    Texto  = f"{mensagem}".replace('{nome}',str(contato)).replace('{cpf}',str(cpf)).replace('{valor}',str(valor))
    UrlTexto = urllib.parse.quote(Texto)
    driverWhats.get(f'https://web.whatsapp.com/send?phone={telefone}&text={UrlTexto}')

    while True:
        print('Carregando...')
        try:
            time.sleep(5)
            WebDriverWait(driverWhats,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="side"]')))
        except: 
            #atualizar a página e tentar novamente
            driverWhats.get(f'https://web.whatsapp.com')
            time.sleep(5)
            driverWhats.get(f'https://web.whatsapp.com/send?phone={telefone}&text={UrlTexto}')
            print(f'https://web.whatsapp.com/send?phone={telefone}&text={UrlTexto}')
        else: break
    return enviar_mensagem_selenium(Texto, cpf, telefone)

def enviar_mensagem_selenium(texto, cpf, telefone):
    #remover o 55 do começo do telefone
    telefone = telefone[2:]
    time.sleep(3)
    #verifica se a mensagem aparece no campo de mensagens quando o link wa.me for acessado

    if Acha('//div[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]', 'verificar'):
        try: print(f'Enviar Mensagem: {texto}')
        except: pass
        driverWhats.find_element(By.XPATH,'//div[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]').send_keys(Keys.ENTER)
        time.sleep(4)

        # Acha('//div[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]', texto)
        # driverWhats.find_element(By.XPATH,'//div[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]').send_keys(Keys.ENTER)
        print(f'Enviado para {telefone}')

        #Gerar uma planilha, com a relação dos cpfs e numeros enviados com sucesso.
        #verificar se existe uma planilha, senão criar
        #se existir, ler a planilha, adicionar o cpf e numero enviado, e salvar a planilha

        
        return True
    else:
        print(f'Não foi possivel enviar mensagem para {telefone}')
        return False

def Acha(xpath, funcao):

    try:
        return Executa(xpath, funcao)
    except:

        return False

def Executa(XPTH,funcao):
    if funcao == 'clica': 
        if Acha(XPTH,'verificar'):
            driverWhats.find_element(By.XPATH,XPTH).click()
    elif funcao =='espera':
        print('Carregado')
    elif funcao == 'limpar': driverWhats.find_element(By.XPATH,XPTH).clear()
    elif funcao == 'valor': return driverWhats.find_element(By.XPATH,XPTH).text
    elif funcao =='valorinput': return driverWhats.find_element(By.XPATH,XPTH).get_attribute("value")
    elif funcao == 'verificar': 
        #Tenta 10 vezes encontrar senão return false
        for i in range(3):
            print(f'Tentando {i+1} vez Verificar {XPTH}')
            try: driverWhats.find_element(By.XPATH,XPTH)
            except: time.sleep(1)
            else: return True
        return False
    else:
        print(f'digitado {funcao}')
        driverWhats.find_element(By.XPATH,XPTH).send_keys(funcao)

def LerFila():
    global pasta
    global TempoMaximoDisparo

    TempoMaximoDisparo = int(TempoMaximo.get())
    #selecionar uma pasta para salvar os arquivos
    pasta = filedialog.askdirectory()
    user_data_dir = f"Disparo_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    options.add_argument(f"--user-data-dir={os.path.abspath(user_data_dir)}")

    #criar um arquivo txt para salvar a mensagem
    with open(f'{pasta}/mensagem.txt', 'w', encoding='utf-8') as f:
        f.write(CaixaDeMensagem.get("1.0", tk.END))
    PlanilhaSucedidos = fr'{pasta}/Sucedidos.xlsx'
    if os.path.exists(PlanilhaSucedidos):
        Sucedidos = pd.read_excel(PlanilhaSucedidos)
    else:
        Sucedidos = pd.DataFrame({'Telefone': '','CPF': ''}, index=[0])
    PlanilhaFalha = fr'{pasta}/Falha.xlsx'
    if os.path.exists(PlanilhaFalha):
        Falha = pd.read_excel(PlanilhaFalha)
    else:
        Falha = pd.DataFrame({'CPF': '','Telefone1': '','Telefone2': '','Telefone3': ''}, index=[0])
    Repeticao = 0
    while Repeticao < TempoMaximoDisparo:
        Fila = PlanilhaFila
        if len(Fila) == 0:
            print('Fila Vazia')
            Repeticao += 1
        else:
            Repeticao = 0
            enviados = 0
            
            for i in range(len(Fila.copy())):
                if enviados == 30:
                    break
                row = Fila[i]
                nomecliente, cpf, telefone1, telefone2, telefone3, valor = row
                telefones = (f"55{telefone1}",f"55{telefone2}",f"55{telefone3}")
                #verificar se os telefones são 550 e ignorar, os demais verificar se envia_mensagem retorna True, se retornar False enviar_mensagem no numero seguinte
                for telefone in telefones:
                    telefone = telefone.replace('.0','').replace(' ','').replace('-','').replace('(','').replace(')','').replace('\n','')
                    if len(telefone) > 13: telefone = telefone[:13]
                    if telefone != "550" and telefone != "55nan":
                        print(f'Verificando {telefone}')
                        if enviar_mensagem(telefone, nomecliente,cpf, valor):
                            enviados += 1
                            Sucedidos = pd.concat([Sucedidos, pd.DataFrame({'Telefone': telefone,'CPF': cpf}, index=[0])])
                            break
                else:
                    #Gerar uma planilha, com a relação dos cpfs e numeros que não foram enviados com sucesso.
                    #verificar se existe uma planilha, senão criar
                    #se existir, ler a planilha, adicionar o cpf e numero enviado, e salvar a planilha
                    Falha = pd.concat([Falha, pd.DataFrame({'CPF': cpf,'Telefone1': telefone1,'Telefone2': telefone2,'Telefone3': telefone3}, index=[0])])
                    Falha.to_excel(PlanilhaFalha, index=False)
            break
            

        time.sleep(5)
    messagebox.showinfo('Fim', 'Fim do Disparo')



inicializacao()
