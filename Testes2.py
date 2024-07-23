#from Funcoes_Cadastro_Empresas import *

#Indo para área de trabalho remota
# pat.moveTo(1425,1044) 
# pat.click()
# time.sleep(2)

from tkinter import Tk, Button

confirma = 2
print(confirma)

def janelaConfirma():
    def botaoConfirmado():
        global confirma
        confirma = 1
        janelaConf.destroy()

    def botaoRepete():
        global confirma
        confirma = 0
        janelaConf.destroy()
    
    janelaConf = Tk()
    janelaConf.title('Deu certo?')
    janelaConf.geometry('350x200')
    botaoConf = Button(janelaConf,text='Exito', command=botaoConfirmado)
    botaoConf.place(x=50,y=50)
    botaoRep = Button(janelaConf, text='Repete processo', command=botaoRepete)
    botaoRep.place(x=100,y=50)
    janelaConf.mainloop()

janelaConfirma()
print(confirma)