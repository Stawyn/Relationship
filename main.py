# -*- coding: latin-1 -*-

# Imports
import pyodbc
import pandas as pd
import time
import warnings
# BAIXAR JUNTO pip install pyodbc
# BAIXAR JUNTO pip install pandas
# BAIXAR JUNTO pip install openpyxl

#----------Começo Inicio----------
warnings.filterwarnings("ignore")  # Suprimir avisos irrelevantes

# Inicialização de variáveis
awa = []
asn = 0-1
contador = 0

# Verificar arquivo de entrada
name = input('Qual é o nome do arquivo Excel que deseja converter? ')
try:
    with open(name + '.xlsx', 'r'):
        pass
except FileNotFoundError:
    print('Não foi possível encontrar o arquivo especificado')
    time.sleep(15)
    exit()

#----------Fim Inicio----------

#----------Começo Função----------
# Funcao LOOP
# Nao alterar
def lopa(linhas, variavel, tabela, nome, codigo, tipo):

    # Variaveis
    global awa, comando, alt, myresult
    global asn
    global contador
    global contescola
    global contcurso
    global contoferta
    global contingresso
    global contturno
    global contmunicipio
    global contuf


    for linha in linhas:
        # Dicionario para substituir
        substituicoes = {'CEPTI': 'CENTRO DE EDUCAÇÃO PROFISSIONAL EM TECNOLOGIA DE INFORMAÇÃO',
                         'CETEP': 'CENTRO DE EDUCAÇÃO TECNOLÓGICA E PROFISSIONALIZANTE',
                         'CVT': 'CENTRO VOCACIONAL TECNOLÓGICO',
                         'EEEF': 'ESCOLA ESTADUAL DE ENSINO FUNDAMENTAL',
                         'ETE': 'ESCOLA TÉCNICA ESTADUAL',
                         'Conc. Ext./Subsequente': 'Concomitância Externa',
                         'Prova': 'Processo Seletivo',
                         'Sorteio/V.O.': 'Sorteio',
                         'Manhã': 'Matutino',
                         'Tarde': 'Vespertino',
                         'Noite': 'Noturno',
                         'Rio de Janeiro Zona Norte': 'Rio de Janeiro',
                         'Rio de Janeiro Zona Sul': 'Rio de Janeiro',
                         '"': '',
                         "'": ''}
        for chave, valor in substituicoes.items():
            linha = linha.replace(chave, valor)
        linha =linha.strip()
        print(hiring['UF candidato'].columns)
        if tipo == 'normal':
            comando = f"""select {codigo} from {tabela} where {tabela}.{nome} like '%{linha}%' COLLATE Latin1_general_CI_AI 
            ORDER BY
              CASE
                WHEN {tabela}.{codigo} LIKE '{linha}' COLLATE Latin1_general_CI_AI THEN 1 
                WHEN {tabela}.{codigo} LIKE '{linha}%' COLLATE Latin1_general_CI_AI THEN 2
                WHEN {tabela}.{codigo} LIKE '%{linha}' COLLATE Latin1_general_CI_AI THEN 4
                ELSE 3
            END"""
        elif tipo == 'municipio':
            comando = f"""select {codigo} from {tabela} where {tabela}.{nome} like '%{linha}%' COLLATE Latin1_general_CI_AI  and UF_COD = (select UF_COD from GER_UF where UF_SIGLA = '{hiring['UF candidato'][asn]}')
            ORDER BY
              CASE
                WHEN {tabela}.{codigo} LIKE '{linha}' COLLATE Latin1_general_CI_AI THEN 1 
                WHEN {tabela}.{codigo} LIKE '{linha}%' COLLATE Latin1_general_CI_AI THEN 2
                WHEN {tabela}.{codigo} LIKE '%{linha}' COLLATE Latin1_general_CI_AI THEN 4
                ELSE 3
            END"""
        try:
            cursor.execute(comando)
            myresult = cursor.fetchall()
            if len(myresult) > 1 and linha not in awa:
                comando2 = f"select {codigo}, {tabela}.{nome} from {tabela} where {tabela}.{nome} like '%{linha}%' COLLATE Latin1_general_CI_AI"
                cursor2.execute(comando2)
                myresult2 = cursor2.fetchall()
                print(f'Parece que o {linha} tem mais registros.\n')
                for i in myresult2:
                    print(i)
                alt = str(input('\nQual será o codigo dele? (nenhum = em branco): '))
                awa.append(linha)
            cursor.execute(comando)
            myresult = cursor.fetchone()
            if myresult == None:
                # print(f'\033[91m{linha}\033[m')
                if awa == linha:
                    variavel.append(alt)
                else:
                    variavel.append(linha)
            else:
                # print(f'\033[92m{myresult[0]}\033[m')
                if awa == linha:
                    variavel.append(alt)
                else:
                    variavel.append(myresult[0])
                    contador += 1
        except:
            # print(f'\033[91;9m{linha}\033[m')
            if awa == linha:
                variavel.append(alt)
            else:
                variavel.append(myresult[0])
    awa = []
    asn += 1
    if tabela == "EDU_ESCOLA":
        contescola = contador
        contador = 0
    if tabela == "EDU_CURSO":
        contcurso = contador
        contador = 0
    if tabela == "EDU_FORMA_OFERTA":
        contoferta = contador
        contador = 0
    if tabela == "EDU_FORMA_INGRESSO":
        contingresso = contador
        contador = 0
    if tabela == "EDU_TURNO":
        contturno = contador
        contador = 0
    if tabela == "GER_MUNICIPIO":
        contmunicipio = contador
        contador = 0
    if tabela == "GER_UF":
        contuf = contador
        contador = 0

    return

#----------Fim Função----------

#----------Começo SQL----------
# DATABASE
# Alteravel
dados_conexao = (
    "Driver={SQL Server};" # Tipo de sql
    "Server=187.108.197.64, 1433;" # IP e porta do servidor
    "Database=Faetec;" # Banco
    'UID=sa;' # Nome do usuario
    'PWD=M1234567890-=m;' # Senha do usuario
)
conexao = pyodbc.connect(dados_conexao)
cursor = conexao.cursor()
cursor2 = conexao.cursor()
print('Conectado aguarde...')

#----------Fim SQL----------

#----------Começo Excel----------
df = pd.read_excel(name + '.xlsx', sheet_name='Dados')
hiring = pd.DataFrame(df)

#criar variavel caso uma nova coluna seja incluida
escola = []
curso = []
oferta = []
ingresso = []
turno = []
sexo = []
nota = []
bairro = nota
cidade = []
uf = []
situacao = nota
notafinal = nota
notaconceito = nota
vaga = nota
raca = nota
sorteio = nota
fase = nota
ano = []

#----------Fim Excel----------

#----------Começo Alteraveis----------

# Linha 1 - Titulos da planilha
# Formato da planilha de saida
# Mantenha EM ORDEM sequencial
# Ao alterar o titulo voce deve alterar os dados inseridos em ordem
columns = ["ESC_COD", "UE", "CUR_COD", "CURSO", "ANO", "FOR_COD", "FORMA OFERTA", "ING_COD", "FORMA DE INGRESSO",
           "EDU_TURNO", "TURNO", "POSICAO", "Nº INSC", "MAT", "DATA NASC", "CANDIDATO", "RG", "CPF", "SEXO_MFN", "SEXO",
           "Telefone Resd.", "Telefone Cel.", "Numero", "CEP", "Endereço", "BAI_COD", "Bairro", "E-mail", "Complemento",
           "MUN_COD", "cidade", "UF _COD", "UF", "NOME DA MAE", "NOTA", "ST_COD", "Situaçao", "NOTA FINAL",
           "NOTACONCEITO", "VAGA_COD", "VAGA", "RAC_COD", "COR DA PELE", "SORTEIO", "FASE", "ANO_CONCURSO"]

# Converção dos dados usando função

# linhas = dados do excel
# variavel = variavel lista para armazenar os dados
# tabela = tabela que vai ser utilizada
# nome = nome que vai ser comparado para achar o codigo
# codigo = o codigo em que sera convertido
# tipo = diferencia normal de municipio
# lopa(hiring['TITULO DA COLUNA QUE IRA PUXAR OS DADOS'],
# Variavel (deve ser criada para guardar os dados tipo "computados = []",
# "EDU_ESCOLA" (nome do banco),
# "ESC_NOME_COMPLETO" (coluna sql nome do dado),
# "ESC_COD" (coluna sql codigo do dado),
# "normal (tipo de dado normal,municipio)")

lopa(hiring['UE'], escola, "EDU_ESCOLA", "ESC_NOME_COMPLETO", "ESC_COD", "normal")
lopa(hiring['Curso'], curso, "EDU_CURSO", "CUR_NOME_REDUZIDO", "CUR_COD", "normal")
lopa(hiring['Forma de Organização'], oferta, "EDU_FORMA_OFERTA", "FOR_NOME", "FOR_COD", "normal")
lopa(hiring['Ingresso'], ingresso, "EDU_FORMA_INGRESSO", "ING_DESCRICAO", "ING_COD", "normal")
lopa(hiring['Turno'], turno, "EDU_TURNO", "TUR_NOME", "TUR_COD", "normal")
for sex in hiring['Sexo']:
    sexo.append(sex.replace('Masculino', 'M').replace('Feminino', 'F'))
# lopa(hiring['Bairro'], bairro, "GER_BAIRRO", "BAI_NOME", "BAI_COD", "normal")
lopa(hiring['Cidade candidato'], cidade, "GER_MUNICIPIO", "MUN_NOME", "MUN_COD", "municipio")
lopa(hiring['UF candidato'], uf, "GER_UF", "UF_SIGLA", "UF_COD", "normal")

for anos in hiring['Ano']:
    ano.append(str(anos).split('.')[0])

for notas in hiring['Cota']:
    nota.append(str(notas).replace(str(notas), ''))

# Criação da nova planilha
# ao criar uma nova coluna voce deve adicionar onde o dado sera inserido
# EM ORDEM
df = pd.DataFrame(list(
    zip(escola, hiring['UE'], curso, hiring['Curso'], ano, oferta, hiring['Forma de Organização'], ingresso,
        hiring['Ingresso'], turno, hiring['Turno'], hiring['Posição'], hiring['Nº INSC'], hiring['MAT DRE'],
        hiring['Data Nasc'], hiring['Nome'], hiring['RG'], hiring['CPF'], sexo, hiring['Sexo'],
        hiring['Telefone Resd.'], hiring['Telefone Cel.'], hiring['Número'], hiring['CEP'], hiring['Endereço'], bairro,
        hiring['Bairro'], hiring['E-mail'], hiring['Complemento'], cidade, hiring['Cidade candidato'], uf,
        hiring['UF candidato'], hiring['Nome Mãe'], nota, situacao, situacao, notafinal, notaconceito, vaga, vaga, raca,
        raca, sorteio, fase, ano)), columns=columns)
df.to_excel(name + ' - Convertido.xlsx', index=False)

#----------Fim Alteraveis----------

#----------Começo Estatisticas----------
print('\n')
print('Terminado arquivo', name + ' - Convertido.xlsx', 'criado')
# Resumo da conversão
print('-' * 40)
print('|{:>35}|'.format('RESUMO DA CONVERSÃO'))
print('-' * 40)

print('|{:>25}: {:<10}|'.format('Arquivo convertido', name + ' - Convertido.xlsx'))
print('|{:>25}: {:<10}|'.format('Total de linhas', len(df.index)))
print('-' * 40)

conttotal = contescola + contcurso + contoferta + contingresso + contturno + contmunicipio + contuf
contoriginal = (len(hiring['UE']) + len(hiring['Curso']) + len(hiring['Forma de Organização']) + len(hiring['Ingresso']) + len(hiring['Turno']) + len(hiring['Cidade candidato']) + len(hiring['UF candidato']))

print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Escola', '', contescola, len(hiring['UE']), (contescola/len(hiring['UE'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Curso', '', contcurso, len(hiring['Curso']), (contcurso/len(hiring['Curso'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Oferta', '', contoferta, len(hiring['Forma de Organização']), (contoferta/len(hiring['Forma de Organização'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Ingresso', '', contingresso, len(hiring['Ingresso']), (contingresso/len(hiring['Ingresso'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Turno', '', contturno, len(hiring['Turno']), (contturno/len(hiring['Turno'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Município', '', contmunicipio, len(hiring['Cidade candidato']), (contmunicipio/len(hiring['Cidade candidato'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('UF', '', contuf, len(hiring['UF candidato']), (contuf/len(hiring['UF candidato'])*100)))
print('|{:>20} {:>3}: {:>3} de {:>3} ({:>6.2f}%)|'.format('Total de sucessos', '', conttotal, contoriginal, (conttotal/contoriginal)*100))
print('-' * 40)

# Taxa de acerto
print('|{:>25}: {:<10.2f}%|'.format('Taxa de acerto', (conttotal/contoriginal)*100))
time.sleep(30)

#----------Fim Estatisticas----------

