import os
import subprocess

### Cria o arquivo de pesquisa
def pesquisar():
    nif = input('Informe o NIF da pessoa: ')
    comando = ('Import-Module activedirectory \n')
    comando1 = ('(Get-ADUser -Identity ' + nif + ' -Properties MemberOf | Select-Object MemberOf).MemberOf > C:\\temp\\user.txt')
    with open('C:\\temp\\user.ps1', 'w', newline="\n", encoding='utf8') as dados:
        dados.write(comando)
        dados.write(comando1)

### Lê o arquivo .txt que esta na pasta temp
def ler():
    global Conteudo
    arquivo = open('C:\\temp\\user.txt', 'r', newline="\n", encoding='utf16')
    conteudo = arquivo.readlines()
    arquivo.close()

### Cria um novo arquivo .csv com os grupos separados
def gravar():
    global Conteudo, ConteudoMod

    for linha in range(len(Conteudo)):
        nova_linha = Conteudo[linha].replace('=', '\t')
        ConteudoMod.append(nova_linha.replace(',', '\t'))

    gravar_novo_arquivo = open('C:\\temp\\novo.csv', 'w', newline='\n', encoding='utf16')
    gravar_novo_arquivo.writelines(ConteudoMod)
    gravar_novo_arquivo.close()

### Chama o PS pelo cmd
def powershell():
    subprocess.call(['powershell', '-Command', 'C:\\temp\\user.ps1'], shell = True)

### Chama o CSV pelo cmd
def csv():
    subprocess.call(['start', 'C:\\temp\\novo.csv'], shell = True)

### Remove o arquivo anterior
if os.path.isfile('C:\\temp\\novo.csv'):
    os.remove('C:\\temp\\novo.csv')


### MAIN
Conteudo = []
ConteudoMod = []

pesquisar()
powershell()
ler()
gravar()
csv()
