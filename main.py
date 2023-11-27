import os
import pandas as pd
import sys
from ctypes import * # Importa toda a lib, está escrito de forma implícita.
from io import StringIO

#### INÍCIO do programa base do PubLog (NÃO MEXER) ###

# As seguintes variáveis e constantes são necessárias para o acesso ao banco de
# dados do PubLog
MAX_SIZE = 4096  # data & error buffer size in bytes
PATH_TO_DLL = "publog\TOOLS\MS12\DecompDl64.dll" # Caminho relativo
PATH_TO_FEDLOG = "publog" # Caminho relativo
dll = CDLL(PATH_TO_DLL) # Carrega a DLL DecompDl64.dll na variável dll.
path = c_char_p(bytes(PATH_TO_FEDLOG, encoding='utf-8')) # A função obtem um 
# array de bytes usando a codificação especificada, no qual c_char_p é um 
# ponteiro C char* para string PATH_TO_FEDLOG que aceita apenas um endereço 
# inteiro ou um objeto em bytes.
data = create_string_buffer(MAX_SIZE) #Cria um buffer de tamanho MAX_SIZE bytes.
error = create_string_buffer(MAX_SIZE)
length = c_int(MAX_SIZE)

if(dll.IMDConnectDLL(path)):
    print("Invalid Path\n")
    sys.exit
    print("Sample search using the return buffer:\n")
#### FIM do programa base do PubLog (NÃO MEXER) ####

def consulta_management_padrao(string_niin):
    """
    A função recebe NIIN em string.
    A função retorna uma lista que contém uma string
    
    - lista_para_retorno = lista para armazenar o resultado da busca no PUBLOG
    - consulta = comando SQL para buscar os vários dados do NIIN, é necessário
    receber o NIIN para fazer a busca, por isso há a concatenação da variável 
    string_niin
    - comando_sql = variável consulta codigicada para acessar o banco do PUBLOG
    - matches = executa a consulta própriamente dita
    - data_convertida = o dado consultado decodificado   
    """
    
    lista_para_retorno = []
    consulta = ("select EFFECTIVE_DATE, MOE, AAC, SOS, UI, UNIT_PRICE, QUP, SLC ")
    consulta2 = ("from V_FLIS_MANAGEMENT where NIIN='" + string_niin + "'")
    comando_sql = (consulta + consulta2).encode('utf-8')
    matches = dll.IMDSqlDLL(comando_sql, data, length)
    data_convertida = data.value.decode('utf-8')
    if data_convertida == '':
        return False
    else:
        lista_para_retorno.extend([data_convertida])
        #return(lista_para_retorno)
        return(data_convertida)


def consulta_management_future(string_niin):
    """
    A função recebe NIIN em string.
    A função retorna uma lista que contém uma string
    
    - lista_para_retorno = lista para armazenar o resultado da busca no PUBLOG
    - consulta = comando SQL para buscar os vários dados do NIIN, é necessário
    receber o NIIN para fazer a busca, por isso há a concatenação da variável 
    string_niin
    - comando_sql = variável consulta codigicada para acessar o banco do PUBLOG
    - matches = executa a consulta própriamente dita
    - data_convertida = o dado consultado decodificado   
    """
    
    lista_para_retorno = []
    consulta = ("select EFFECTIVE_DATE, MOE, AAC, SOS, UI, UNIT_PRICE, QUP, SLC ")
    consulta2 = ("from V_FLIS_MANAGEMENT_FUTURE where NIIN='" + string_niin + "'")
    comando_sql = (consulta + consulta2).encode('utf-8')
    matches = dll.IMDSqlDLL(comando_sql, data, length)
    data_convertida = data.value.decode('utf-8')
    if data_convertida == '':
        return False
    else:
        lista_para_retorno.extend([data_convertida])
        return(lista_para_retorno)


def filtrar_quantidade_digitos_string(i):
    """
    A função recebe uma string na variável i.
    A função retorna o número dentro da string com 9 dígitos.
    A ideia é transformar um NSN em um NIIN removendo os 4 primeiros dígitos e
    completar com o dígito 0 caso a string tenha menos do que 9 dígitos. O 
    motivo disso é possíveis erros dentro da planilha por se tratar do dígito
    0 à esquerda.
    
    - a variavel i representa a string oriunda da planilha a ser filtrada
    """
        
    if len(i) < 9:
        i = i.rjust(9, '0')
        
    while len(i) > 9:
        i = i.replace(i[0], '', 1)
    return i    


def verificar_aac(string_niin):
    """
    A função recebe um NIIN em formato string e retorna um Boolean. 
    A função retorna um Bool.
    O retorno True acusa a presença de AAC não desejados.
    
    - AAC_nao_aceitaveis = Lista de AAC que não interessam ao usuário
    - consulta = Comando SQL que recebe o NIIN em formato de string
    - conmando_sql = Conversão da string para bytes, codificação do comando
    - matches = O retorno da quantidade de itens buscados no PubLog
    - data_convertida = O retorno da informação do PubLog decodificada
    - data_temporaria = data_convertida sem caracteres especiais 
    - data_filtrada = data_temporaria sem os três primeiros caracteres (AAC)
    """
    
    try:
        aac_nao_aceitaveis = ['F', 'L', 'P', 'V', 'X', 'Y', 'T']
        consulta = (
                    "select AAC, from V_FLIS_MANAGEMENT where NIIN='" + string_niin + "'"
                    )
        comando_sql = consulta.encode('utf-8')
        matches = dll.IMDSqlDLL(comando_sql, data, length)
        data_convertida = data.value.decode('utf-8')
        data_temporaria = "".join(c for c in data_convertida if c.isalnum())
        data_filtrada = data_temporaria[3:]
        print(data_filtrada) # REMOVER ESSA LINHA APÓS TESTES
        return (set(data_filtrada) <= set(aac_nao_aceitaveis))
        # retorna True se todos os aac baterem com os aac não aceitaveis
                       
    except: # Ainda são desconhecidos possíveis erros específicos para tratar
        print(f'Erro na função verificar_aac. Foi usado o NIIN {string_niin}')


def verificar_ui_box_pg(string_niin):
    """
    A função recebe o NIIN em string.
    A função retorna o Bool True se o NIIN tiver PG ou BX
    
    - string_niin = NIIN em formato string recebida
    - ui_pg = string dos caracteres que se deseja verificar o UI
    - ui_bx = string dos caracteres que se deseja verificar o UI
    - consulta = comando SQL para buscar o UI do NIIN, é necessário receber o NIIN
    para fazer a busca, por isso há a concatenação da variável string_niin
    - comando_sql = variável consulta codigicada para acessar o banco do PUBLOG
    - matches = executa a consulta própriamente dita
    - data_convertida = o dado consultado decodificado
    - data_temporaria = data_convertida sem caractéres que não sejam números e 
    letras
    - data_filtrada = data_convertida retidado os caractéres UI da string,
    sobrando PG ou BX no dado filtrado
    - ui_pg_presente = lógica para verificar se o dado do PUBLOG é UI que se quer
    verificar
    - ui_bx_presente = lógica para verificar se o dado do PUBLOG é UI que se quer
    verificar
    """
    
    try:
        ui_pg = 'PG' 
        ui_bx = 'BX'
        consulta = (
                    "select UI, from V_FLIS_MANAGEMENT where NIIN='" + string_niin + "'"
                    )
        comando_sql = consulta.encode('utf-8')
        matches = dll.IMDSqlDLL(comando_sql, data, length)
        data_convertida = data.value.decode('utf-8')
        data_temporaria = "".join(c for c in data_convertida if c.isalnum())
        data_filtrada = data_temporaria[2:]
        ui_pg_presente = (set(data_filtrada) <= set(ui_pg))
        ui_bx_presente = (set(data_filtrada) <= set(ui_bx))        
        print(set(data_filtrada) <= set(ui_pg)) # REMOVER ESSA LINHA APÓS TESTES
        if ui_bx_presente or ui_pg_presente:
            return True    
                       
    except: # Ainda são desconhecidos possíveis erros específicos para tratar
        print(f'Erro na função verificar_ui_box_pg. Foi usado o NIIN {string_niin}.')
        

def consultar_quantidade_box_pg(string_niin):
    """
    A função recebe NIIN em string.
    A função retorna ... a função não retorna nada
    
    - consulta = comando SQL para buscar o PHRASE_STATEMENT do NIIN, é necessário
    receber o NIIN
    para fazer a busca, por isso há a concatenação da variável string_niin
    - comando_sql = variável consulta codigicada para acessar o banco do PUBLOG
    - matches = executa a consulta própriamente dita
    - data_convertida = o dado consultado decodificado    
    - ...

    """
        
    consulta = (
                "select PHRASE_STATEMENT, from V_FLIS_PHRASE where NIIN='" + string_niin + "'"
                )
    comando_sql = consulta.encode('utf-8')
    matches = dll.IMDSqlDLL(comando_sql, data, length)
    data_convertida = data.value.decode('utf-8')
    print('*****CONSULTANDO BOX***** ' + string_niin) # REMOVER ESSA LINHA APÓS TESTES
    print([data_convertida])
    return data_convertida


def main(lista_de_niin):
    for string in lista_de_niin: 
        string = filtrar_quantidade_digitos_string(str(string))
        
        if verificar_aac(string): # Verifica o aac
            print(f'o NIIN {string}, tem o aac não desejável')
            # Função pra jogar informação na planilhas
            continue
        
        if verificar_ui_box_pg(string):
            consultar_quantidade_box_pg(string)
        
        if consulta_management_future(string):
            a = (consulta_management_future(string)) 
            print(a) # REMOVER ESSA LINHA APÓS TESTES
            print('*****FUTURO***** ' + string) # REMOVER ESSA LINHA APÓS TESTES
            continue
            # Função pra jogar informação na planilha       
        
        if consulta_management_padrao(string):
            a = (consulta_management_padrao(string)) 
            print(a) # REMOVER ESSA LINHA APÓS TESTES
            print('*****PADRAO***** ' + string) # REMOVER ESSA LINHA APÓS TESTES
            # Função pra jogar informação na planilha  
            df = pd.read_csv(StringIO(a), sep='|')
            print(df)
            
os.system('cls') # limpar console
print("Programa de Pesquisa de Preço por NIIN.\n\n") # REMOVER ESSA LINHA APÓS TESTES

# Ler o arquivo csv
df = pd.read_csv("arquivo_niin.csv")
col_list = df.NSN.values.tolist()

lista_teste = ['15600022901']
main(col_list)
#main(lista_teste)


# df = pd.read_csv(StringIO(data), sep='|')
# df.to_csv()
