# Progama feito por Luiz Felipe Carneiro De Oliveira 
# email: luizfelipe.c.d.o@gmail.com      tel.: (64) 99931-8800

from Funcoes_Cadastro_Empresas import *

################################### Encerrando área de trabalho remota ########################################################

fechandoAreaRemota()

############################## Começando automação ###################################################################

grupo = janelaInput() # Abre janela para coletar o grupo de peças a ser corrigido nas 3 lojas

pat.alert('O código vai começar, não toque no teclado ou no mouse enquanto isso!') # Manda alerta na tela com o texto 

abrindoAreaRemota() # Abre area de trabalho remota, claro kkkk
fechandoApps() # Fecha no máximo 4 coisas abertas na area remota


abrindoCigoELogin() # Abrindo sistema cigo e fazendo o login

janelaConfirma(13,fechandoAreaRemota,abrindoAreaRemota,fechandoApps,abrindoCigoELogin)

pesquisandoNoPesquisaItem(grupo, pesquisa=True) # Pesquisar grupo de peças

janelaConfirma(0,[pesquisandoNoPesquisaItem,grupo,True])

exportandoParaExcel() # Exportar dados para formato excel

copiandoArquivoExcel() # Abrindo local do arquivo exportado e copiando arquivo

colandoExcelEmTrabalhoLoja(grupo) # Colando arquivo na pasta TRABALHO LOJA e renomeando

janelaConfirma(0, exportandoParaExcel, copiandoArquivoExcel, [colandoExcelEmTrabalhoLoja,grupo])

########################## Análise de dados utilizando pandas #################################################################

arquivo = ('C:/Users/luizf/OneDrive/Desktop/TRABALHO LOJA/' + 'PRODUTOS ' + grupo + ' - ' + dataParaSalvamento + '.xls')

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

janelaConfirma(5, fechandoAreaRemota, [time.sleep, 3], [pat.keyDown, 'alt'], [pat.press, 'f4'], [pat.keyUp, 'alt'])

pat.PAUSE = 0.75
pat.MINIMUM_DURATION = 0.35

igualandoCadastros(falta_wag,falta_elm,falta_soc)

janelaConfirma(0, [igualandoCadastros,falta_wag,falta_elm,falta_soc])