#from Funcoes_Cadastro_Empresas import *

#Indo para área de trabalho remota
# pat.moveTo(1425,1044) 
# pat.click()
# time.sleep(2)

def som(a, b):
    return a + b

lista = [som, 2, 3]

def impr(lis):
        tupla = tuple(lis[1:])
        return lis[0]

print(impr(lista))

# [print(i) for i in lista[1:]]