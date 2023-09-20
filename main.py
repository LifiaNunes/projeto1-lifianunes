import xmltodict
import os
import pandas as pd

def pegar_infos(nome_arquivo, valores):
    #print (f"pegou as informações {nome_arquivo}") -- mostra quais arquivos xml foram lidos
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml: # -- realiza leitura dos arquivos "r" no diretorio /nfs + arquivo xml
        dic_arquivo = xmltodict.parse(arquivo_xml) # -- transforma os dados em xml para pyton e armazena no dic_arquivo, usa o parse da lib

        if "NFe" in dic_arquivo:
          infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"] # -- feito esse if pois há divergencia no nome da raiz entre os arquivos xml analisados
        numero_da_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf["emit"]["xNome"]
        nome_cliente = infos_nf["dest"]["xNome"]
        endereco = infos_nf["dest"]["enderDest"]
        if "vol" in infos_nf["transp"]:
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Não informado" # -- feito para também tratar divergencia no nome raiz "pesoB" existentes entres os arquivos XML analisados
        
        endereco_formatado = ( # --- faz a tratativa do nome dos dados advindos do XML: 
            f'Logradouro: {endereco["xLgr"]}, '
            f'Número: {endereco["nro"]}, '
            f'Complemento: {endereco.get("xCpl", "")}, '
            f'Bairro: {endereco["xBairro"]}, '
            f'Município: {endereco["xMun"]}, '
            f'UF: {endereco["UF"]}, '
            f'CEP: {endereco["CEP"]}, '
            f'País: {endereco["xPais"]}, '
            f'Fone: {endereco.get("fone", "Não informado")}'
        )
        valores.append([numero_da_nota, empresa_emissora, nome_cliente, endereco_formatado, peso]) # -- criado a variavel para armazenar os registros, o append vem no panda para armazenar as informações de acordo com as posições informadas
    
lista_arquivos = os.listdir("nfs") # -- lib os, faz uma listagem dos dados e armazena na lista_arquivos

colunas = ["numero_da_nota", "empresa_emissora", "nome_cliente", "endereco", "peso"] # --- colunas criadas no excel quando exportar os dados XML solicitados acima
valores = [] # --- aqui serão disponibilizados os registros, em ordem feita acima no valores.append (lib pd)

for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores) # -- percorre o script, busca os dados e traz os registros "valores"

tabela = pd.DataFrame(columns=colunas, data=valores) # -- criação da tabela usando a lib pd, chamando em pd.DataFrame, criando as colunas (columns) e dados (data)
#print (tabela) -- para mostrar a tabela


#criação de arquivo em Excel com os dados recebidos do XML 
tabela.to_excel("NotasFiscais1.xlsx", index=False)         
