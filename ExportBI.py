import pandas as pd
from pymongo import MongoClient
import os

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"
NOME_ARQUIVO_SAIDA = "dados_para_powerbi.csv"
# --------------------

def exportar_dados_para_csv():
    """
    Conecta ao MongoDB, busca todos os pedidos, "achata" os dados dos itens
    e exporta tudo para um único arquivo CSV, pronto para o Power BI.
    """
    print("Iniciando o exportador de dados para o Power BI...")
    
    try:
        print("Conectando ao MongoDB...")
        client = MongoClient(MONGO_CONNECTION_STRING)
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        
        # Busca todos os documentos da coleção
        dados = list(collection.find({}))
        
        if not dados:
            print("Aviso: Nenhum dado encontrado no banco de dados para exportar.")
            return
            
        print(f"Encontrados {len(dados)} pedidos. Processando e achatando os dados...")
        
        # Lista para armazenar os dados "achatados"
        linhas_finais = []
        
        # Itera sobre cada documento de pedido
        for pedido in dados:
            # Itera sobre cada item dentro do pedido
            for item in pedido.get('itens', []):
                nova_linha = {
                    # Informações do Pedido (repetidas para cada item)
                    'ID_Pedido_Filial': pedido.get('_id'),
                    'Numero_PV': pedido.get('numero_pv'),
                    'Filial_Codigo': pedido.get('filial_codigo'),
                    'Filial_Nome': pedido.get('filial_nome'),
                    'Parceiro': pedido.get('parceiro'),
                    'Data_Emissao': pedido.get('emissao'),
                    'Vendedor': pedido.get('vendedor'),
                    'Condicao_Pagamento': pedido.get('condicao_pagamento'),
                    'Valor_Total_Pedido': pedido.get('valor_total_pedido'),
                    
                    # Informações do Item
                    'Cod_Produto': item.get('cod_produto'),
                    'Descricao_Produto': item.get('descricao'),
                    'Quantidade_Item': item.get('quantidade'),
                    'Valor_Unitario_Item': item.get('unitario'),
                    'Valor_Total_Item': item.get('total_item')
                }
                linhas_finais.append(nova_linha)

        if not linhas_finais:
            print("Nenhum item encontrado nos pedidos para exportar.")
            return

        # Cria o DataFrame final com os dados achatados
        df_final = pd.DataFrame(linhas_finais)
        
        # Garante que a coluna de data esteja no formato correto
        df_final['Data_Emissao'] = pd.to_datetime(df_final['Data_Emissao'])
        
        # Salva o arquivo CSV
        caminho_saida = os.path.join(os.getcwd(), NOME_ARQUIVO_SAIDA)
        df_final.to_csv(caminho_saida, index=False, sep=';', decimal=',', encoding='utf-8-sig')
        
        print("\n--- SUCESSO! ---")
        print(f"Os dados foram exportados com sucesso para o arquivo:")
        print(caminho_saida)
        
    except Exception as e:
        print(f"\n--- ERRO ---")
        print(f"Ocorreu um erro durante a exportação: {e}")

if __name__ == '__main__':
    exportar_dados_para_csv()
