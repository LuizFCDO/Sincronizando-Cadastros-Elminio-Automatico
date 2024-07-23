import pyautogui as pat
import time
from tkinter import Tk, Button # biblioteca para janela de comunicação
from tkinter import simpledialog # biblioteca para janela de comunicação que recebe informação
from os import kill # biblioteca para comandos do sistema
import psutil as ps # biblioteca para processos do sistema
import signal # biblioteca para encerrar processos
import pandas as pd
from IPython.display import display # biblioteca com display para data fremes do pandas

#import PosicoesEmpresa # para importar as posições de ponteiro no computador da empresa
import PosicoesParticular # para importar as posições de ponteiro no computador particular

def pidAreaRemota():
    for proc in ps.process_iter(): # Armazenando pid da área de trabalho remota no pidArea
        info = proc.as_dict(attrs=['pid', 'name'])
        if info['name'] == 'mstsc.exe':
            pidArea = int(info['pid'])
            return pidArea
    return 0

def finalizar_processo(pid): # Função para encerrar processos com base no pid
    kill(pid, signal.SIGTERM)

def fechandoAreaRemota():
    pidArea = pidAreaRemota() # Obtendo o pid da atea de trabalho temota e armazenando em pidArea
    if(pidArea != 0): # Encerrando área de trabalho remota
        finalizar_processo(pidArea)

def janelaInput():
    janela = Tk()   
    janela.title('Janela de entrada')
    janela.geometry('350x200')
    grupo = simpledialog.askstring("Grupo a ser verificado", "Digite o código do grupo de peças a ser verificado:", parent=janela)
    janela.destroy()
    janela.mainloop()
    return grupo

def janelaConfirma(tempo,*comandos):
    def botaoConfirmado():
        janelaConf.destroy()

    def botaoRepete():
        for i in comandos:
            if type(i)==list:
                if i.len() == 2:
                    i[0](i[1])
                elif i.len() == 3:
                    i[0](i[1], i[2])
                elif i.len() == 4:
                    i[0](i[1], i[2], i[3])
                elif i.len() == 5:
                    i[0](i[1], i[2], i[3], i[4])
            else:
                i()
    
    janelaConf = Tk()
    janelaConf.title('Deu certo?')
    janelaConf.geometry('350x200')
    botaoConf = Button(janelaConf,text='Exito', command=botaoConfirmado)
    botaoConf.place(x=50,y=50)
    botaoRep = Button(janelaConf, text='Repete processo', command=botaoRepete)
    botaoRep.place(x=100,y=50)
    janelaConf.mainloop()
    time.sleep(tempo)

pat.PAUSE = 0.5
pat.MINIMUM_DURATION = 0.25

dataParaSalvamento = time.strftime("%d %m %Y", time.localtime())

empr = {
    'elm' : 'ELMINIO',
    'soc' : 'SOCORRO',
    'wag' : 'WAGNER'
}

def abrindoAreaRemota():
    # Abrindo a área de trabalho remota
    pat.press('win', interval=1)
    pat.write('conex~ao de ´area de trabalho',interval=0.2)
    pat.press('enter', presses=2, interval=1)
    time.sleep(10)

def fechandoApps():
    # Fechando possíveis janelas abertas na área remota
    pat.keyDown('alt') # Equivalente a pat.hotkey('alt', 'f4')
    pat.press('f4', presses=4, interval=0.2)
    pat.keyUp('alt')
    pat.press('esc')

def abrindoCigoELogin():
    time.sleep(1)  
    pat.moveTo(cord['cigo atalho'])
    pat.click(clicks=2)
    time.sleep(8)
    pat.write("21", interval=0.3)
    pat.press('tab')
    pat.write("k123d", interval=0.3)
    pat.press('enter')

def pesquisandoNoPesquisaItem(grupo, pesquisa=True):
    pat.keyDown('alt') 
    pat.press('b')
    pat.keyUp('alt')
    time.sleep(13)
    pat.moveTo(cord['pesq item 4 cat'])
    pat.click()
    pat.moveTo(cord['cod. ref pesq item 4 cat'])
    pat.click()
    if pesquisa == True:
        pat.write(grupo, interval=0.25)
        pat.moveTo(cord['lupa pesq item'])
        pat.click()

def exportandoParaExcel():
    pat.moveTo(cord['itens pesq item'])
    time.sleep(7)
    pat.click(button='right')
    pat.press('x')
    time.sleep(12)

def copiandoArquivoExcel():
    pat.press('win')
    pat.write('documentos')
    pat.press('enter')
    pat.press('o')
    pat.press('enter')
    pat.press('e')
    pat.press('enter')
    pat.press('c')
    pat.keyDown('ctrl')
    pat.press('c')
    pat.keyUp('ctrl')

def colandoExcelEmTrabalhoLoja(grupo):
    pat.moveTo(cord['min area rem'])
    pat.click()
    pat.press('win')
    pat.write('TRABALHO LOJA')
    pat.press('enter')
    time.sleep(2)
    pat.press('down')
    pat.hotkey('ctrl', 'v')
    time.sleep(10)
    pat.press('f2')
    time.sleep(3)
    pat.write('PRODUTOS ' + grupo + ' - ' + dataParaSalvamento)
    pat.press('enter')

def abrindoProduto():
    pat.keyDown('alt') 
    pat.press('t')
    pat.keyUp('alt')
    time.sleep(13)

def pesquisandoEmProduto(ref,pesquisa):
    pat.moveTo(cord['prod 4 cat'])
    pat.click()
    pat.moveTo(cord['cod. ref prod 4 cat'])
    pat.click()
    pat.write(ref, interval=0.25)
    pat.moveTo(cord['lupa prod']) # movendo para lupa
    if pesquisa == True:
        pat.click()
        time.sleep(1)

def cadastroEmpresaProduto(empresa):
    # copiando grupo de empresa já cadastrada
    pat.moveTo(cord['empr prod']) # selecionando empresa produto
    pat.click()
    pat.moveTo(cord['alt empr prod']) # selecionando alteração empresa produto 
    pat.click()
    time.sleep(4)
    pat.moveTo(cord['grup alt empr prod']) # selecionando grupo, dentro de alteração empresa produto
    pat.click()
    time.sleep(3)
    pat.keyDown('ctrl') # copiando grupo do produto para novo cadastro
    pat.press('c', presses=3, interval=0.3)
    pat.keyUp('ctrl')
    pat.moveTo(cord['cancel alt empr prod']) # selecionando cancelar, dentro de alteração empresa produto
    time.sleep(2)
    pat.click()
    time.sleep(2)
    pat.press('n')
    time.sleep(1)

    # adicionando nova empresa
    pat.moveTo(cord['adic empr prod']) # selecionando adicionar empresa produto
    pat.click()
    time.sleep(6)
    pat.write(empresa) # digitando a empresa para cadastro do produto
    pat.press('enter')
    pat.press('enter')
    time.sleep(9)
    pat.keyDown('ctrl') # colando o grupo do produto no novo cadastro
    pat.press('v', presses=3, interval=0.3)
    pat.keyUp('ctrl')
    pat.sleep(3)
    pat.press('enter', presses=4, interval=1)
    pat.write('mercadoria') # escolhendo o tipo do produto
    pat.press('enter')
    time.sleep(3)
    pat.press('f5') # salvando novo cadastro
    time.sleep(6)

def igualandoCadastros(falta_wag, falta_elm, falta_soc):
    for cod in falta_wag:
        pesquisandoEmProduto(cod,pesquisa=True)
        cadastroEmpresaProduto('wag')
    for cod in falta_elm:
        pesquisandoEmProduto(cod,pesquisa=True)
        cadastroEmpresaProduto('elm')
    for cod in falta_soc:
        pesquisandoEmProduto(cod,pesquisa=True)
        cadastroEmpresaProduto('soc')