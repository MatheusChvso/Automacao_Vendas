from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"
# --------------------

def verificar_duplicatas():
    """Conecta ao MongoDB e procura por pedidos duplicados."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING, serverSelectionTimeoutMS=5000)
        # Testa a conexão
        client.admin.command('ping')
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        print(f"Conectado com sucesso ao banco '{MONGO_DATABASE}' e coleção '{MONGO_COLLECTION}'.")
    except ConnectionFailure as e:
        print(f"Não foi possível conectar ao MongoDB: {e}")
        return

    # --- Verificação 1: Duplicatas Exatas (mesmo _id) ---
    # Isto é tecnicamente impossível no MongoDB, mas é uma boa verificação inicial.
    pipeline_id_exato = [
        {"$group": {"_id": "$_id", "count": {"$sum": 1}}},
        {"$match": {"count": {"$gt": 1}}}
    ]
    print("\n--- Verificando duplicatas de _id exato (deve retornar vazio)...")
    duplicatas_exatas = list(collection.aggregate(pipeline_id_exato))
    if not duplicatas_exatas:
        print(">> OK: Nenhuma duplicata de _id exato encontrada, como esperado.")
    else:
        print(f"!! ATENÇÃO: Encontradas {len(duplicatas_exatas)} duplicatas de _id. Isso é muito incomum.")
        print(duplicatas_exatas)

    # --- Verificação 2: Duplicatas por Lógica (mesmo pedido, filial com maiúscula/minúscula diferente) ---
    # Esta é a verificação mais importante, que encontra o problema de "filial_RJ" vs "filial_rj".
    pipeline_logica = [
        {
            "$group": {
                "_id": {
                    "numero_pv": "$numero_pv",
                    # Converte o código da filial para minúsculas para agrupar "SS" e "ss" juntos
                    "filial_normalizada": {"$toLower": "$filial_codigo"}
                },
                "documentos": {"$push": {"_id": "$_id", "filial_original": "$filial_codigo", "data_carga": "$data_carga"}},
                "count": {"$sum": 1}
            }
        },
        {"$match": {"count": {"$gt": 1}}}
    ]
    
    print("\n--- Verificando duplicatas lógicas (ex: 'RJ' vs 'rj')...")
    duplicatas_logicas = list(collection.aggregate(pipeline_logica))

    if not duplicatas_logicas:
        print(">> OK: Nenhuma duplicata lógica encontrada.")
        print("\nConclusão: Seu banco de dados está íntegro e sem duplicatas!")
    else:
        print(f"\n!! ATENÇÃO: Foram encontrados {len(duplicatas_logicas)} grupos de pedidos duplicados logicamente!")
        print("Isso geralmente acontece por salvar arquivos com nomes de filial diferentes (ex: filial_RJ.xlsx e filial_rj.xlsx).")
        print("Abaixo estão os grupos de duplicatas encontrados:")
        for grupo in duplicatas_logicas:
            print("-" * 30)
            print(f"Pedido Nº: {grupo['_id']['numero_pv']}, Filial: {grupo['_id']['filial_normalizada']}")
            for doc in grupo['documentos']:
                print(f"  - _id: {doc['_id']}, Filial Original: '{doc['filial_original']}', Carregado em: {doc['data_carga']}")
    
    client.close()


if __name__ == '__main__':
    verificar_duplicatas()