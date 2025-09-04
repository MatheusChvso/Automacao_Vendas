import pandas as pd
from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"
# --------------------

def analisar_periodo_dados():
    """Conecta ao MongoDB e analisa o intervalo de datas dos registros."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING, serverSelectionTimeoutMS=5000)
        client.admin.command('ping')
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        print(f"Conectado com sucesso ao banco '{MONGO_DATABASE}'.")
    except ConnectionFailure as e:
        print(f"Não foi possível conectar ao MongoDB: {e}")
        return

    print("\nAnalisando o período dos dados na coleção 'pedidos'...")
    
    # Busca todas as datas de emissão
    todas_as_datas = list(collection.find({}, {"emissao": 1, "_id": 0}))

    if not todas_as_datas:
        print(">> Nenhum documento encontrado na coleção.")
        client.close()
        return
        
    df = pd.DataFrame(todas_as_datas)
    df['emissao'] = pd.to_datetime(df['emissao'])
    df.dropna(subset=['emissao'], inplace=True)

    if df.empty:
        print(">> Nenhum documento com data de emissão válida foi encontrado.")
        client.close()
        return

    # Encontra a data mais antiga e a mais recente
    data_mais_antiga = df['emissao'].min()
    data_mais_recente = df['emissao'].max()
    
    print("\n--- RESULTADO DA ANÁLISE ---")
    print(f"Total de registros com data válida: {len(df)}")
    print(f"Venda mais ANTIGA registrada: {data_mais_antiga.strftime('%d/%m/%Y')}")
    print(f"Venda mais RECENTE registrada: {data_mais_recente.strftime('%d/%m/%Y')}")
    
    # Conta os registros por ano
    print("\nDistribuição de registros por ano:")
    contagem_por_ano = df['emissao'].dt.year.value_counts().sort_index()
    print(contagem_por_ano)
    
    print("\n--------------------------")
    print("\nConclusão: O script de relatório está procurando por dados a partir de 01/01/2025.")
    print("Se a maioria dos seus dados for de anos anteriores, é normal que os gráficos apareçam vazios.")

    client.close()

if __name__ == '__main__':
    analisar_periodo_dados()