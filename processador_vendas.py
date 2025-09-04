import pandas as pd
import os
import shutil
from pymongo import MongoClient
from datetime import datetime

# --- CONFIGURAÇÕES - AJUSTE ESTA SEÇÃO ---

# 1. Caminhos das pastas
PASTA_ENTRADA = 'Para_Processar'
PASTA_ARQUIVO = 'Processados'
PASTA_ERRO = 'Erro'

# 2. Mapeamento de Filiais (NOVA SEÇÃO)
#    Mapeia o código do arquivo para o nome completo da filial.
MAPA_FILIAIS = {
    "SS": "Solution",
    "Va": "Vale Aço",
    "SZM": "Zona da Mata",
    "RJ": "Rio de Janeiro"
}

# 3. Configurações do MongoDB
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"

# 4. Nomes das colunas do arquivo exportado
COL_NUMERO_PV = "Numero PV"
COL_EMISSAO = "Emissão"
COL_PARCEIRO = "Parceiro"
COL_VENDEDOR = "Vendedor Externo"
COL_PRODUTO = "Cod. Produto"
COL_PRODUTO_DESC = "Descrição"
COL_QTD = "Quantidade"
COL_UNITARIO = "Unitário"
COL_TOTAL_ITEM = "Total"
COL_COND_PAGTO = "CONDPAGTO"
# --- FIM DAS CONFIGURAÇÕES ---


def conectar_mongodb():
    """Estabelece a conexão com o MongoDB e retorna a coleção."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING)
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        client.admin.command('ping')
        print("Conexão com o MongoDB bem-sucedida!")
        return collection
    except Exception as e:
        print(f"Erro ao conectar ao MongoDB: {e}")
        return None

def transformar_em_pedidos(df):
    """Agrupa as linhas de itens em documentos de pedido únicos."""
    print("Transformando dados de itens para o formato de pedido...")
    
    pedidos_formatados = []

    # Agrupa o DataFrame pelo número do pedido e pelas informações da filial
    for (numero_pv, filial_cod), grupo in df.groupby([COL_NUMERO_PV, 'filial_codigo']):
        
        info_pedido = grupo.iloc[0]
        
        itens_do_pedido = []
        for _, item_row in grupo.iterrows():
            item = {
                "cod_produto": item_row[COL_PRODUTO],
                "descricao": item_row[COL_PRODUTO_DESC],
                "quantidade": item_row[COL_QTD],
                "unitario": item_row[COL_UNITARIO],
                "total_item": item_row[COL_TOTAL_ITEM]
            }
            itens_do_pedido.append(item)
            
        pedido_doc = {
            "_id": f"{int(numero_pv)}_{filial_cod}", # Chave única com o código da filial
            "numero_pv": int(numero_pv),
            "filial_codigo": filial_cod,
            "filial_nome": info_pedido['filial_nome'], # Campo com o nome completo
            "parceiro": info_pedido[COL_PARCEIRO],
            "emissao": info_pedido[COL_EMISSAO],
            "vendedor": info_pedido[COL_VENDEDOR],
            "condicao_pagamento": info_pedido[COL_COND_PAGTO],
            "valor_total_pedido": grupo[COL_TOTAL_ITEM].sum(),
            "itens": itens_do_pedido,
            "data_carga": datetime.now()
        }
        
        pedidos_formatados.append(pedido_doc)
        
    print(f"Transformação concluída. {len(pedidos_formatados)} pedidos únicos encontrados.")
    return pedidos_formatados


def processar_arquivos():
    """Função principal que orquestra todo o processo."""
    print("Iniciando o processador de vendas...")
    
    collection = conectar_mongodb()
    if collection is None:
        return

    arquivos_para_processar = [f for f in os.listdir(PASTA_ENTRADA) if f.endswith(('.csv', '.xlsx'))]
    
    if not arquivos_para_processar:
        print("Nenhum arquivo encontrado para processar.")
        return

    lista_dfs = []
    for arquivo in arquivos_para_processar:
        caminho_arquivo = os.path.join(PASTA_ENTRADA, arquivo)
        try:
            print(f"Processando arquivo: {arquivo}")
            if arquivo.endswith('.csv'):
                df_temp = pd.read_csv(caminho_arquivo, sep=';', decimal=',')
            else:
                df_temp = pd.read_excel(caminho_arquivo)
    
           
            codigo_filial_encontrado = None
            for codigo in MAPA_FILIAIS.keys():
                if codigo in arquivo:
                    codigo_filial_encontrado = codigo
                    break # Para assim que encontrar o primeiro código
            
            if codigo_filial_encontrado:
                nome_completo_filial = MAPA_FILIAIS[codigo_filial_encontrado]
                df_temp['filial_codigo'] = codigo_filial_encontrado
                df_temp['filial_nome'] = nome_completo_filial
                print(f"  -> Filial '{nome_completo_filial}' identificada.")
                lista_dfs.append(df_temp)
            else:
                print(f"  -> AVISO: Nenhuma filial conhecida (SS, Va, SZM, RJ) encontrada no nome do arquivo '{arquivo}'. Arquivo ignorado.")
           
        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")
            shutil.move(caminho_arquivo, os.path.join(PASTA_ERRO, arquivo))
            
    if not lista_dfs:
        print("Nenhum arquivo foi lido com sucesso.")
        return
        
    df_consolidado = pd.concat(lista_dfs, ignore_index=True)
    print("Colunas encontradas no arquivo:", df_consolidado.columns)

    df_consolidado[COL_EMISSAO] = pd.to_datetime(df_consolidado[COL_EMISSAO], errors='coerce')
    df_consolidado[COL_TOTAL_ITEM] = pd.to_numeric(df_consolidado[COL_TOTAL_ITEM], errors='coerce')
    df_consolidado.dropna(subset=[COL_NUMERO_PV, COL_EMISSAO, COL_TOTAL_ITEM], inplace=True)
    
    pedidos_para_processar = transformar_em_pedidos(df_consolidado)

    if pedidos_para_processar:
        chaves_do_lote = [p['_id'] for p in pedidos_para_processar]
        documentos_existentes = collection.find({'_id': {'$in': chaves_do_lote}}, {'_id': 1})
        chaves_existentes = {doc['_id'] for doc in documentos_existentes}
        print(f"Encontradas {len(chaves_existentes)} chaves de pedidos já existentes no banco.")

        pedidos_para_inserir = [p for p in pedidos_para_processar if p['_id'] not in chaves_existentes]

        if pedidos_para_inserir:
            print(f"Inserindo {len(pedidos_para_inserir)} novos pedidos no MongoDB...")
            collection.insert_many(pedidos_para_inserir)
            print("Novos pedidos inseridos com sucesso.")
        else:
            print("Nenhum pedido novo para inserir.")
    
    for arquivo in arquivos_para_processar:
        caminho_origem = os.path.join(PASTA_ENTRADA, arquivo)
        if os.path.exists(caminho_origem):
             shutil.move(caminho_origem, os.path.join(PASTA_ARQUIVO, f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_{arquivo}"))

    print("Processo concluído!")

if __name__ == "__main__":
    processar_arquivos()