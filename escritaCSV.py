import os, shutil, subprocess

### Cria o arquivo de pesquisa
def pesquisa():
    nif = input('Informe o NIF da pessoa: ')
    comando = ('Import-Module activedirectory \n')
    comando1 = ('(Get-ADUser -Identity ' +nif+ ' -Properties MemberOf | Select-Object MemberOf).MemberOf > C:\\temp\\user.txt')
    with open('C:\\temp\\user.ps1', 'w', newline="\n", encoding='utf8') as dados:
        dados.write(comando)
        dados.write(comando1)

### Lê o arquivo .txt que esta na pasta temp
def ler():
    global conteudo
    lerArquivo = open('C:\\temp\\user.txt', 'r', newline="\n", encoding='utf16')
    conteudo = lerArquivo.readlines()
    #print(conteudo)
    print(len(conteudo))
    lerArquivo.close()

### Cria um novo arquivo .csv com os grupos separados
def gravar():
    global conteudo, conteudoMod

    for linha in range(len(conteudo)):
        novaLinha = conteudo[linha].replace('=', '\t')
        conteudoMod.append(novaLinha.replace(',', '\t'))

    gravaNovoArquivo = open('C:\\temp\\novo.csv', 'w', newline='\n', encoding='utf16')
    gravaNovoArquivo.writelines(conteudoMod)
    gravaNovoArquivo.close()

### Chama o PS pelo cmd
def chamarPS():
    subprocess.call(['powershell', '-Command', 'C:\\temp\\user.ps1'], shell = True)

### Chama o CSV pelo cmd
def chamarCsv():
    subprocess.call(['start', 'C:\\temp\\novo.csv'], shell = True)

### Remove o arquivo anterior
if os.path.isfile('C:\\temp\\novo.csv'):
    os.remove('C:\\temp\\novo.csv')


### MAIN
conteudo = []
conteudoMod = []

pesquisa()
chamarPS()
ler()
gravar()
chamarCsv()
