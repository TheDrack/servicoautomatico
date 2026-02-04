from tkinter import *
from tkinter import mainloop
import pyautogui
import time
import speech_recognition as sr
import pyttsx3
import os
import sys
import pandas as pd

audio = sr.Recognizer()
maquina = pyttsx3.init()


def avisar(fala):
    print(fala)
    falar(fala)
    pyautogui.alert(fala)

def falar (fala):
    maquina.say(fala)
    maquina.runAndWait()

def digitar (texto):
    pyautogui.write(texto)

def aperta (botao):
    pyautogui.press(botao)

def restart():
    python = sys.executable
    os.execl(python, python, * sys.argv)


def Ligar_microfone ():

    with sr.Microphone() as fonte:

        while True:    
            audio.adjust_for_ambient_noise(fonte)
            voz = audio.listen(fonte)
            comando = audio.recognize_google(voz, language='pt-BR')
            comando = comando.lower()

            if 'cancelar' in comando:
                falar('cancelado ação')
                janela.destroy
                inicio()
            elif 'voltar' in comando:
                pyautogui.hotkey('shift','tab')
                comando = comando.replace('voltar', '')
                return(comando)
            elif 'fechar' in comando:
                sys.exit()
            else :
                return (comando)

def chamarAXerife ():
    os.startfile(os.path.dirname(os.path.realpath(__file__)) + r"\assistente.pyw")

def pedircomando():
    falar('Diga um comando')
    ordem = Ligar_microfone()
    comandos(ordem)

tempoDeEspera = 7.5
def localizanatela(imagem):
    caminho = os.path.dirname(os.path.realpath(__file__)) + r'\imagens'
    arquivo = imagem
    k = 0
    n = tempoDeEspera
    os.chdir(caminho)

    while True:
        #Procura a imagem
        local = pyautogui.locateCenterOnScreen(arquivo)

        #Se imagem for localizada 
        if local != None:
            pyautogui.moveTo(local)
            print(f'Imagem {imagem} localizada na posição: {local}')
            return True

        #Após n tentativas o programa encerra
        if k >= n:
            print(f'Imagem {imagem} não localizada')
            return False

        #Aguarda um pouco para tentar novamente
        time.sleep(0.25)
        k += 1

def excel (LocalExcel):
    df = pd.read_excel(LocalExcel)
    out = df.to_numpy().tolist()
    Tupla = [tuple(elt) for elt in out]
    return Tupla

def abrirgaveta ():
    pyautogui.hotkey('ctrl','shift','g')

def FazerRequisicaoPT1 ():
    mic = Ligar_microfone
    AbrirRequisicao()
    EscolherCentroDeCusto(mic)
    descritivoRequisicao(mic)
    AnotacaoRequisicao(mic)
    FazerRequisicaoPT2()
def FazerRequisicaoPT2 ():
    mic = Ligar_microfone
    digitar(Cod4rMaterial(mic))
    QuantMaterial(mic)
    AlgoMais(mic)

def AbrirRequisicao ():
    falar('Irá inicializar uma requisição automatizada, não clique em nada')
    pyautogui.PAUSE = 0.4
    if localizanatela('botaoALMOX.PNG') == True:
        pyautogui.click()
        pyautogui.hotkey('win','up')
        pyautogui.leftClick(200,50)
        pyautogui.leftClick(770,90)
        pyautogui.doubleClick(100,175)
        pyautogui.PAUSE = 0.6
        aperta(['tab','tab','tab'])
        pyautogui.PAUSE = 0.4
        digitar('1')
        aperta(['enter','enter'])
    else :
        pyautogui.alert("botão não localizado")

def EscolherCentroDeCusto (CC):
    falar('Qual centro de custo será o destinatário?')
    TuplaDeCC = excel(r"ServicoAutomatico\Lista de CC.xlsx")
    for codigo,escrito in TuplaDeCC:
        codigo = str(codigo)
        if escrito in CC:       
            CC = codigo
    digitar(CC)
    aperta('enter')
    aperta('enter')


def descritivoRequisicao (desc):
    falar('Diga o descritivo')
    desc = desc.upper()
    digitar(desc)
    aperta('enter')


def AnotacaoRequisicao(AC):
    falar('Está aos cuidados de qual solicitante?')

    AC = AC.upper()
    digitar('A/C ' + AC)
    aperta('enter')



def Cod4rMaterial (escolher):
    falar('Diga o material')
    Material = escolher
    TuplaDeMateriais = excel(os.path.dirname(os.path.realpath(__file__)) + r"\Lista de Materiais.xlsx")
    for codigo,material in TuplaDeMateriais:
        codigo = str(codigo)
        if material in Material:
            Material = '0'+ codigo
    return(Material)


def QuantMaterial (Quantidade):
    falar('Fale a quantidade')
    if 'pular' in Quantidade:
        prox()
        pyautogui.hotkey('shift','tab')
    else:
        aperta('enter')
    digitar(Quantidade)


def AlgoMais(resposta):
    falar('Algo mais?')
    if 'sim' in resposta:
        FazerRequisicaoPT2()
    elif 'não'  in resposta:
        falar('Pronto')
        restart()
    else :
        falar('Não entendi')

def prox ():
    aperta('enter')
    digitar('1')
    aperta('enter')

def FazerRequisicaoSulfite ():
    pyautogui.PAUSE = 0.4
    if localizanatela('botaoALMOX.PNG'):
        pyautogui.click()
        pyautogui.leftClick(200,50)
        pyautogui.leftClick(770,90)
        pyautogui.doubleClick(100,175)
        aperta(['tab','tab','tab'])
        digitar('1')
        pyautogui.press('enter',presses=6)
        digitar('0301603043')
        aperta('enter')
        digitar('0,5')
        pyautogui.alert('Automatização concluida, continue manualmente')
    else :
        pyautogui.alert('botão de Almoxarifado 4R não encontrado')

def AtualizarInventario ():
    pyautogui.PAUSE = 0.4
    if localizanatela('botaoALMOX.PNG'):
        pyautogui.click()
        pyautogui.leftClick(200,50)
        pyautogui.doubleClick(100,270)
        pyautogui.doubleClick(160,165)
        pyautogui.leftClick(155,175)
        pyautogui.leftClick(155,175)
        pyautogui.leftClick(585,310)
        pyautogui.leftClick(585,400)
        pyautogui.leftClick(585,480)
        pyautogui.leftClick(170,505)
        pyautogui.leftClick(170,600)
        if localizanatela('botaoPDF.PNG'):
            pyautogui.click()
            pyautogui.leftClick(20,30)
            aperta('home')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('enter')
            aperta('enter')
            pyautogui.hotkey('shift','tab')
            pyautogui.hotkey('shift','tab')
            pyautogui.hotkey('shift','tab')
            aperta('home')
            aperta('a')
            time.sleep(1)
            aperta('enter')
            aperta('tab')
            aperta('tab')
            aperta('tab')
            digitar('Material 10\Rafael\Consumo.inventario\Dados\inventario.xls')
            aperta('enter')
            aperta('enter')
            pyautogui.leftClick(2000,10)
            pyautogui.alert('Automatização concluida, continue manualmente')
        else: 
            pyautogui.alert('botão PDF não encontrado')
    else: 
        pyautogui.alert('botão ALMOX não encontrado')

def imprimirAjusteDeEstoque ():
    pyautogui.PAUSE = 0.4
    if localizanatela('botaoALMOX.PNG'):
        pyautogui.click()
        pyautogui.leftClick(200,50)
        pyautogui.doubleClick(100,270)
        pyautogui.doubleClick(160,165)
        pyautogui.leftClick(155,175)
        pyautogui.leftClick(155,175)
        pyautogui.leftClick(585,310)
        pyautogui.leftClick(585,400)
        pyautogui.leftClick(585,480)
        pyautogui.leftClick(585,552)
        pyautogui.leftClick(170,600)
        if localizanatela('botaoPDF.PNG'):
            pyautogui.click()
            pyautogui.leftClick(20,30)
            aperta('home')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('down')
            aperta('enter')
            aperta('enter')
            pyautogui.hotkey('shift','tab')
            pyautogui.hotkey('shift','tab')
            pyautogui.hotkey('shift','tab')
            aperta('home')
            aperta('e')
            time.sleep(1)
            aperta('enter')
            aperta('tab')
            digitar('d')
            aperta('enter')
            aperta('tab')
            aperta('tab')
            digitar('GitHub\ServicoAutomatico\AjusteDeEstoque')
            aperta('enter')
            aperta('enter')
            pyautogui.moveTo(1,1)
            if localizanatela('botaoIMPRESSAO.PNG'):
                pyautogui.click()
            else:
                pyautogui.alert('botão para impressão não localizado')
        else: 
            pyautogui.alert('botão PDF não encontrado')
    else: 
        pyautogui.alert('botão ALMOX não encontrado')

def lancarAjusteDeEstoque ():
    AbrirRequisicao()
    digitar('451')
    pyautogui.press('enter')
    pyautogui.press('enter')
    descritivoRequisicao('Ajuste de estoque apos conferencia de inventario')
    AnotacaoRequisicao('Rafael')

    PastaBase = os.path.dirname(os.path.realpath(__file__))

    TuplaDeMateriais = excel(PastaBase + r"\ajustedeestoque.xls")
    for codigo,Material,Unid,Mov,Qntd,Conferido in TuplaDeMateriais:
        codigo = str(codigo).replace('.','')
        if Conferido != '':
            if Conferido < Qntd:
                digitar(Cod4rMaterial(codigo))
                avisar(str(Conferido) + ' de ' + str(Qntd))
                Lancar = int(Conferido - Qntd)
                QuantMaterial(str(Lancar))
                aperta('enter')
    avisar("Lançamento de ajuste de estoque concluido")

def ImprimirBalancete ():
    pyautogui.PAUSE = 0.4
    if localizanatela('botaoALMOX.PNG'):
        pyautogui.click()
        pyautogui.leftClick(200,50)
        pyautogui.doubleClick(100,270)
        pyautogui.doubleClick(160,155)
        pyautogui.leftClick(240,200)
        digitar('1')
        pyautogui.leftClick(585,320)
        pyautogui.leftClick(585,440)
        pyautogui.leftClick(585,565)
        pyautogui.leftClick(400,150)
        pyautogui.alert('Automatização concluida, continue manualmente')
    else:
        pyautogui.alert("botão não localizado")

def AbrirPlanilha ():
    pyautogui.PAUSE = 0.4
    pyautogui.hotkey('ctrl', 'shift', 'i')
    pyautogui.alert('ESPERE. Clique em Ok quando a planilha abrir para não dar erro')
    aperta('enter')
    time.sleep(2)
    pyautogui.hotkey('win', 'right')
    pyautogui.doubleClick(1500,270)
    pyautogui.hotkey('ctrl', 'b')
    pyautogui.leftClick(660,1050)
    pyautogui.hotkey('win','left')
    pyautogui.alert('Automatização concluida, continue manualmente')

def AbrirAlmox ():
    pyautogui.PAUSE = 0.8
    pyautogui.alert("O código vai começar. Não use nada do seu computador enquanto o código está rodando")
    pyautogui.hotkey('ctrl', 'shift', 'a')
    localizanatela('login4R.PNG')
    digitar('jesus.anhaia')
    aperta('tab')
    digitar('123456')
    aperta('enter')
    aperta('enter')
    pyautogui.alert('Almoxarifado 4R Aberto, prossiga manualmente')

def VisualizarPasta():

    pyautogui.alert(os.getcwd())
    pyautogui.alert(os.path.dirname(os.path.realpath(__file__)))

def ConsultarEstoque():
    material = Cod4rMaterial(Ligar_microfone)
    TuplaDeMateriais = excel("/inventario.xls")
    Quantidade = None
    for codigo,Material,Unid,Mov,Qntd, in TuplaDeMateriais:
        codigo = str(codigo)
        if material in codigo:
            Quantidade = Qntd
            break
    if Quantidade is None:
        falar('Não foi encontrado')
    return Quantidade


comandos = {"requisição":FazerRequisicaoPT1,"sulfite":FazerRequisicaoSulfite,"planilha":AbrirPlanilha,"inventário":AtualizarInventario,"balancete":ImprimirBalancete,"almoxarifado":AbrirAlmox,"digitar produto":Cod4rMaterial,"digitar quantidade":QuantMaterial,"gaveta":abrirgaveta,"escreva":digitar,"aperte":aperta,"falar":falar,"internet":pyautogui.hotkey('ctrl','shift','c'),"estoque":ConsultarEstoque,"fechar":sys.exit}

janela = Tk()

janela.title("Automatização")
janela.minsize(1500,800)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=0, row=0, padx=5, pady=5)

botao = Button(janela, text="Fazer Requisição", command=FazerRequisicaoPT1)
botao.grid(column=0, row=1, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=1, row=0, padx=5, pady=5)

botao = Button(janela, text="Atualizar Inventario", command=AtualizarInventario)
botao.grid(column=1, row=1, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=1, row=2, padx=5, pady=5)

botao = Button(janela, text="Imprimir Ajuste de Estoque", command=imprimirAjusteDeEstoque)
botao.grid(column=1, row=3, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=1, row=4, padx=5, pady=5)

botao = Button(janela, text="Lançar Ajuste de Estoque", command=lancarAjusteDeEstoque)
botao.grid(column=1, row=5, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=2, row=0, padx=5, pady=5)

botao = Button(janela, text="Imprimir Balancete", command=ImprimirBalancete)
botao.grid(column=2, row=1, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=3, row=0, padx=5, pady=5)

botao = Button(janela, text="Abrir Planilha", command=AbrirPlanilha)
botao.grid(column=3, row=1, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=0, row=2, padx=5, pady=5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=0, row=4, padx=5, pady=5)

botao = Button(janela, text="Fazer Requisição Sulfite", command=FazerRequisicaoSulfite)
botao.grid(column=0, row=5, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=3, row=2, padx=5, pady=5)

botao = Button(janela, text="Abrir Almoxarifado 4R", command=AbrirAlmox)
botao.grid(column=3, row=3, padx= 5, pady= 5)

texto_orientacao = Label(janela, text="Clique para executar a ação automatizada")
texto_orientacao.grid(column=3, row=4, padx=5, pady=5)

botao = Button(janela, text="Abrir Planilha de Gaveta", command=abrirgaveta)
botao.grid(column=3, row=5, padx= 5, pady= 5)

botao = Button(janela, text="Visualizar pasta atual", command=VisualizarPasta)
botao.grid(column=4, row=3, padx= 5, pady= 5)

botao = Button(janela, text="Chamar a Xerife", command=chamarAXerife)
botao.grid(column=4, row=4, padx= 5, pady= 5)

def inicio():
    janela.mainloop()
inicio()