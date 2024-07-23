from Funcoes_Cadastro_Empresas import *

# #Indo para área de trabalho remota
# pat.moveTo(1425,1044) 
# pat.click()
# time.sleep(2)

grupo = '400.'
arquivo = ('C:/Users/rosky/Desktop/Automações Luiz/Planilhas Sincronizando Cadastros/' + 'PRODUTOS ' + grupo + ' - ' + dataParaSalvamento + '.xls')
colandoExcelEmTrabalhoLoja(grupo,arquivo) # Colando arquivo na pasta TRABALHO LOJA e renomeando