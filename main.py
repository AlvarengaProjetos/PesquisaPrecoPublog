import csv
import sys
from ctypes import * # Importa toda a lib, está escrito de forma implícita.


def criar_lista(string_tab):
    """
    Retorna uma lista com sublistas sem Duplicatas de NIIN e FSC
    """
    lista = []
    palavra = ""
    for i in string_tab:
        if i != '<':
            if i == '>':
                lista.append(palavra)
                palavra = ""
            elif i != '\n':
                palavra = palavra + i
    return lista
    

# 1º Etapa: Pesquisar NIIN através do PN e CFF(CAGE_CODE)
def pesquisar_niin(PN, CFF):
    print("> Part Number: "+ PN + " - CFF: " + CFF)
    consulta = (
        "select NIIN from P_PART_PICK where PART_NUMBER='"+ PN +"'"+" AND CAGE_CODE='"+ CFF +"'"
        ) 
    comandoSQL = consulta.encode('utf-8') # Conversão da string para bytes
    matches = dll.IMDSqlDLL(comandoSQL, data, length)

    if matches == 0:
        print("Part Number/CFF não encontrado.\n")
        return 0
    elif matches == 1:
        print("Foi encontrado apenas 1 item.")
    elif matches > 1:
        print("Foram encontrados " + str(matches) + " itens.")
    
    string = (data.value.decode('utf-8'))
    lista = criar_lista(string)
    return lista[1]


# 2º Etapa: Pesquisar Unit_Price utilizando NIIN acrescentando na Lista
def pesquisar_preco_medio(NIIN):
    sniin = str(NIIN)
    consulta = ("select UNIT_PRICE from V_FLIS_MANAGEMENT where NIIN='" + sniin.strip() + "'")
    comandoSQL = consulta.encode('utf-8')  # Conversão da string para bytes
    matches = dll.IMDSqlDLL(comandoSQL, data, length)
    
    if matches == 1:
        print("Apenas um preço foi encontrado para este item.")
    else:
        print("Há " + str(matches) + " preços para este item.")
    
    string_preco_unidade = (data.value.decode('utf-8'))
    lista = criar_lista(string_preco_unidade)
    soma = 0
    for i in lista[1:]:
        soma = soma + float(i)
    media = soma / (len(lista) - 1)
    return media


# Função para retorno de mensagem de quantidade de preços
def mensagem_quantidade_preços(NIIN):
    sniin = str(NIIN)
    consulta = ("select UNIT_PRICE from V_FLIS_MANAGEMENT where NIIN='" + sniin.strip() + "'")
    comandoSQL = consulta.encode('utf-8')  # Conversão da string para bytes
    matches = dll.IMDSqlDLL(comandoSQL, data, length)
    
    if matches == 1:
        return"Apenas um preço foi encontrado para este item."
    else:
        return"Há " + str(matches) + " preços para este item."


#---------------------------------- Programa Principal ---------------------------------

# As seguintes variáveis e constantes são necessárias para o acesso ao 
MAX_SIZE = 4096  # data & error buffer size in bytes
PATH_TO_DLL = "publog\TOOLS\MS12\DecompDl64.dll" # Caminho relativo, presume
PATH_TO_FEDLOG = "publog" # Caminho relativo
dll = CDLL(PATH_TO_DLL) # Carrega a DLL DecompDl64.dll na variável dll.
path = c_char_p(bytes(PATH_TO_FEDLOG, encoding='utf-8')) # A função obtem um array de bytes usando a codificação especificada, no qual c_char_p é um ponteiro C char* para string PATH_TO_FEDLOG que aceita apenas um endereço inteiro ou um objeto em bytes.
data = create_string_buffer(MAX_SIZE) # Cria um buffer de tamanho MAX_SIZE bytes.
#error = create_string_buffer(MAX_SIZE)
length = c_int(MAX_SIZE)

print("Programa de Pesquisa de Preço por PN e CFF.\n\n")

if(dll.IMDConnectDLL(path)):
    print("Invalid Path\n")
    sys.exit
#print("Sample search using the return buffer:\n")

arquivo = open("arquivo.csv")
s_tabela = csv.reader(arquivo, delimiter=';')
tabela = list(s_tabela)

tabela[0].append("NIIN")
tabela[0].append("Preço Médio Pesquisado (U$)")

print(*tabela, sep='\n')


    
tabela[0] = [
        'PN', 'CFF', 'NOMECLATURA', 'UN', 'PRICE UNIT(U$)', 'Retorno da busca', 'Comparação preço'
    ]


#
for indice, elemento in enumerate(tabela[1:]):
    PN = elemento[0]
    CFF = elemento[1]
    NIIN = pesquisar_niin(PN, CFF)

    if NIIN == 0:
        elemento.append('Part Number/CFF não encontrado.')
        elemento.append('Não há preço')

    elif NIIN != 0:
        mensagem_de_precos = mensagem_quantidade_preços(NIIN)
        preco_medio_publog = float(pesquisar_preco_medio(NIIN))
        preco_da_planilha = float(elemento[4])

        print(preco_da_planilha)
        print(preco_medio_publog)


        elemento.append(f'{mensagem_de_precos} Preço médio: ${preco_medio_publog}.')

        if preco_medio_publog > preco_da_planilha:
            resultado = round(((preco_medio_publog - preco_da_planilha) / preco_da_planilha * 100), 2)
            elemento.append(f'Preço médio do PubLog é {resultado}% mais caro.')
        
        elif preco_medio_publog < preco_da_planilha:
            resultado = round(((preco_da_planilha - preco_medio_publog) / preco_medio_publog * 100), 2)
            elemento.append(f'Preço médio do PubLog é {resultado}% mais barato.') 

        else:
            elemento.append(f'O preço médio do PubLog é igual ao preço da planilha.') 
        
        print("Preço médio: $"+ str(preco_medio_publog)+".\n\n")


# CRIAR PLANILHA NO FORMATO LIBRE OFFICE
with open('Preços_PubLog_Comparados_LibreOffice.csv', 'w', newline='') as arquivo_data:
    escritor = csv.writer(arquivo_data)
    escritor.writerows(tabela)

# CRIAR PLANILHA NO FORMATO EXCEL
with open('Preços_PubLog_Comparados_Excel.csv', 'w', newline='') as arquivo_data:
    separador = ['sep=,']

    escritor = csv.writer(arquivo_data)
    escritor.writerow(separador)
    escritor.writerows(tabela)

          
