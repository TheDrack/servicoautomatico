from tkinter import *
from tkinter import mainloop
import pyautogui
import time
import speech_recognition as sr
import pyttsx3
import os
import sys
import pandas as pd
from typing import List, Tuple, Any


audio = sr.Recognizer()
maquina = pyttsx3.init()


def avisar(fala: str) -> None:
    print(fala)
    falar(fala)
    pyautogui.alert(fala)

def falar(fala: str) -> None:
    maquina.say(fala)
    maquina.runAndWait()

def digitar(texto: str) -> None:
    pyautogui.write(texto)

def aperta(botao: str | List[str]) -> None:
    pyautogui.press(botao)

def restart() -> None:
    python = sys.executable
    os.execl(python, python, * sys.argv)


def Ligar_microfone() -> str:

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

def chamarAXerife() -> None:
    os.startfile(os.path.dirname(os.path.realpath(__file__)) + r"\assistente.pyw")

def pedircomando() -> None:
    falar('Diga um comando')
    ordem = Ligar_microfone()
    comandos(ordem)

tempoDeEspera = 7.5


def localizanatela(imagem: str) -> bool:
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

def excel(LocalExcel: str) -> List[Tuple[Any, ...]]:
    df = pd.read_excel(LocalExcel)
    out = df.to_numpy().tolist()
    Tupla = [tuple(elt) for elt in out]
    return Tupla

def abrirgaveta() -> None:
    pyautogui.hotkey('ctrl','shift','g')

def FazerRequisicaoPT1() -> None:
    mic = Ligar_microfone
    AbrirRequisicao()
    EscolherCentroDeCusto(mic)
    descritivoRequisicao(mic)
    AnotacaoRequisicao(mic)
    FazerRequisicaoPT2()
def FazerRequisicaoPT2() -> None:
    mic = Ligar_microfone
    digitar(Cod4rMaterial(mic))
    QuantMaterial(mic)
    AlgoMais(mic)

def AbrirRequisicao() -> None:
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

def EscolherCentroDeCusto(CC: str) -> None:
    falar('Qual centro de custo será o destinatário?')
    TuplaDeCC = excel(r"ServicoAutomatico\Lista de CC.xlsx")
    for codigo,escrito in TuplaDeCC:
        codigo = str(codigo)
        if escrito in CC:       
            CC = codigo
    digitar(CC)
    aperta('enter')
    aperta('enter')


def descritivoRequisicao(desc: str) -> None:
    falar('Diga o descritivo')
    desc = desc.upper()
    digitar(desc)
    aperta('enter')


def AnotacaoRequisicao(AC: str) -> None:
    falar('Está aos cuidados de qual solicitante?')

    AC = AC.upper()
    digitar('A/C ' + AC)
    aperta('enter')



def Cod4rMaterial(escolher: str) -> str:
    falar('Diga o material')
    Material = escolher
    TuplaDeMateriais = excel(os.path.dirname(os.path.realpath(__file__)) + r"\Lista de Materiais.xlsx")
    for codigo,material in TuplaDeMateriais:
        codigo = str(codigo)
        if material in Material:
            Material = '0'+ codigo
    return(Material)


def QuantMaterial(Quantidade: str) -> None:
    falar('Fale a quantidade')
    if 'pular' in Quantidade:
        prox()
        pyautogui.hotkey('shift','tab')
    else:
        aperta('enter')
    digitar(Quantidade)


def AlgoMais(resposta: str) -> None:
    falar('Algo mais?')
    if 'sim' in resposta:
        FazerRequisicaoPT2()
    elif 'não'  in resposta:
        falar('Pronto')
        restart()
    else :
        falar('Não entendi')

def prox() -> None:
    aperta('enter')
    digitar('1')
    aperta('enter')

def FazerRequisicaoSulfite() -> None:
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

def AtualizarInventario() -> None:
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
            digitar(r'Material 10\Rafael\Consumo.inventario\Dados\inventario.xls')
            aperta('enter')
            aperta('enter')
            pyautogui.leftClick(2000,10)
            pyautogui.alert('Automatização concluida, continue manualmente')
        else: 
            pyautogui.alert('botão PDF não encontrado')
    else: 
        pyautogui.alert('botão ALMOX não encontrado')

def imprimirAjusteDeEstoque() -> None:
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
            digitar(r'GitHub\ServicoAutomatico\AjusteDeEstoque')
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

def lancarAjusteDeEstoque() -> None:
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

def ImprimirBalancete() -> None:
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

def AbrirPlanilha() -> None:
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

def AbrirAlmox() -> None:
    pyautogui.PAUSE = 0.8
    pyautogui.alert("O código vai começar. Não use nada do seu computador enquanto o código está rodando")
    pyautogui.hotkey('ctrl', 'shift', 'a')
    localizanatela('login4R.PNG')
    # NOTA: Credenciais devem ser obtidas de variáveis de ambiente
    # por segurança. Hardcoded apenas para compatibilidade com versão histórica.
    usuario = os.getenv('ALMOX_USER', 'jesus.anhaia')
    senha = os.getenv('ALMOX_PASS', '123456')
    digitar(usuario)
    aperta('tab')
    digitar(senha)
    aperta('enter')
    aperta('enter')
    pyautogui.alert('Almoxarifado 4R Aberto, prossiga manualmente')

def VisualizarPasta() -> None:

    pyautogui.alert(os.getcwd())
    pyautogui.alert(os.path.dirname(os.path.realpath(__file__)))

def ConsultarEstoque() -> Any:
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

# Criação dos botões de forma organizada
botoes_config = [
    ("Fazer Requisição", FazerRequisicaoPT1, 0, 1),
    ("Atualizar Inventario", AtualizarInventario, 1, 1),
    ("Imprimir Ajuste de Estoque", imprimirAjusteDeEstoque, 1, 3),
    ("Lançar Ajuste de Estoque", lancarAjusteDeEstoque, 1, 5),
    ("Imprimir Balancete", ImprimirBalancete, 2, 1),
    ("Abrir Planilha", AbrirPlanilha, 3, 1),
    ("Fazer Requisição Sulfite", FazerRequisicaoSulfite, 0, 5),
    ("Abrir Almoxarifado 4R", AbrirAlmox, 3, 3),
    ("Abrir Planilha de Gaveta", abrirgaveta, 3, 5),
    ("Visualizar pasta atual", VisualizarPasta, 4, 3),
    ("Chamar a Xerife", chamarAXerife, 4, 4),
]

for texto, comando, col, row in botoes_config:
    Button(janela, text=texto, command=comando).grid(column=col, row=row, padx=5, pady=5)

def inicio() -> None:
    janela.mainloop()
inicio()