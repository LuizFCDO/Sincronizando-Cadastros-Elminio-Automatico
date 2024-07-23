# Progama feito por Luiz Felipe Carneiro De Oliveira 
# email: luizfelipe.c.d.o@gmail.com      tel.: (64) 99931-8800

from Funcoes_Cadastro_Empresas import *

############################################ Função para confirmar execução correta ###############################################################

confirma = 0

def janelaConfirma():
    def botaoConfirmado():
        global confirma
        confirma = 1
        janelaConf.destroy()
        return confirma

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

################################### Encerrando área de trabalho remota ########################################################

fechandoAreaRemota()

############################## Começando automação ###################################################################

grupo = janelaInput() # Abre janela para coletar o grupo de peças a ser corrigido nas 3 lojas

pat.alert('O código vai começar, não toque no teclado ou no mouse enquanto isso!') # Manda alerta na tela com o texto 

abrindoAreaRemota() # Abre area de trabalho remota, claro kkkk
fechandoApps() # Fecha no máximo 4 coisas abertas na area remota


abrindoCigoELogin() # Abrindo sistema cigo e fazendo o login
time.sleep(5)

# janelaConfirma()
# while confirma == 0:
#     fechandoAreaRemota()
#     abrindoAreaRemota()
#     fechandoApps()
#     abrindoCigoELogin()
#     janelaConfirma()

pesquisandoNoPesquisaItem(grupo, pesquisa=True) # Pesquisar grupo de peças
time.sleep(5)

# janelaConfirma()
# while confirma == 0:
#     pesquisandoNoPesquisaItem(grupo)
#     janelaConfirma()

exportandoParaExcel() # Exportar dados para formato excel

copiandoArquivoExcel() # Abrindo local do arquivo exportado e copiando arquivo

arquivo = ('C:/Users/rosky/Desktop/Automações Luiz/Planilhas Sincronizando Cadastros/' + 'PRODUTOS ' + grupo + ' - ' + dataParaSalvamento + '.xls')
colandoExcelEmTrabalhoLoja(grupo,arquivo) # Colando arquivo na pasta TRABALHO LOJA e renomeando
time.sleep(5)

# janelaConfirma()
# while confirma == 0:
#     abrindoAreaRemota()
#     time.sleep(3)
#     pat.keyDown('alt')
#     pat.press('f4')
#     pat.keyUp('alt')
#     time.sleep(5)
#     exportandoParaExcel()
#     copiandoArquivoExcel()
#     colandoExcelEmTrabalhoLoja()
#     janelaConfirma()

########################## Análise de dados utilizando pandas #################################################################

dados_df = pd.read_excel(arquivo,dtype='str')

dados_df = dados_df[['Cód. ref', 'Empresa']] # Separando só referência e empresa
quant_reg = dados_df['Cód. ref'].value_counts() # Vendo quantas vezes cada referência aparece na tabela (tem que haver 3 ocorrências)
quem_falta = quant_reg.loc[quant_reg < 3] # Verificando todas as referências que não estão nas 3 empresas

falta_elm = [] #lista das referências que faltam na elminio
falta_soc = [] #lista das referências que faltam na socorro
falta_wag = [] #lista das referências que faltam na wagner

for i in quem_falta.index: # preenchendo as listas das peças que faltam em cada empresa
    aux = dados_df.loc[dados_df['Cód. ref'] == i, 'Empresa']
    if(not('ELMINIO' in aux.values)):
        falta_elm.append(i)
    if(not('SOCORRO' in aux.values)):
        falta_soc.append(i)
    if(not('WAGNER' in aux.values)):
        falta_wag.append(i)

################################## Voltando para automação ##################################################################################

fechandoAreaRemota()

abrindoAreaRemota() # Reabrindo área de trabalho remota 
time.sleep(3)

pat.keyDown('alt') # Fechando o explorador de arquivos
pat.press('f4')
pat.keyUp('alt')
time.sleep(5)

abrindoProduto() # Abrindo aba de produto no cigo

# janelaConfirma()
# while confirma == 0:
#     fechandoAreaRemota()
#     abrindoAreaRemota()
#     fechandoApps()
#     abrindoCigoELogin()
#     abrindoProduto()
#     janelaConfirma()

pat.PAUSE = 0.75
pat.MINIMUM_DURATION = 0.35

igualandoCadastros(falta_wag,falta_elm,falta_soc)

pat.alert('Fim dos processos!')