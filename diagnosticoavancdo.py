from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"
# --------------------

def encontrar_duplicatas_logicas():
    """Conecta ao MongoDB e procura por documentos que são funcionalmente idênticos."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING, serverSelectionTimeoutMS=5000)
        client.admin.command('ping')
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        print(f"Conectado com sucesso ao banco '{MONGO_DATABASE}'.")
    except ConnectionFailure as e:
        print(f"Não foi possível conectar ao MongoDB: {e}")
        return

    # Pipeline de agregação para encontrar duplicatas com base nos dados de negócio
    pipeline = [
        {
            # Passo 1: Agrupa por campos que definem uma duplicata lógica
            "$group": {
                "_id": {
                    "numero_pv": "$numero_pv",
                    "parceiro": "$parceiro",
                    "emissao": "$emissao",
                    "valor": "$valor_total_pedido"
                },
                # Coleta os _ids e filial_codigo de todos os documentos no grupo
                "documentos_encontrados": {
                    "$push": {
                        "_id": "$_id",
                        "filial_codigo": "$filial_codigo",
                        "data_carga": "$data_carga"
                    }
                },
                "count": {"$sum": 1}
            }
        },
        {
            # Passo 2: Filtra apenas os grupos que têm mais de um documento (duplicatas)
            "$match": {
                "count": {"$gt": 1}
            }
        }
    ]

    print("\n--- Procurando por duplicatas lógicas avançadas...")
    print("(Pedidos com mesmo número, parceiro, data de emissão e valor total)")
    
    duplicatas = list(collection.aggregate(pipeline))

    if not duplicatas:
        print("\n>> Nenhuma duplicata encontrada com a verificação avançada.")
        print("Isso é muito estranho, dado o aumento nos valores. O problema pode ser outro.")
    else:
        print(f"\n!! ENCONTRADO: {len(duplicatas)} grupos de pedidos duplicados foram identificados!")
        print("Abaixo está o detalhamento:")
        for grupo in duplicatas:
            info = grupo['_id']
            docs = grupo['documentos_encontrados']
            print("-" * 50)
            print(f"Pedido Duplicado: Nº {info['numero_pv']} | Parceiro: {info['parceiro']} | Valor: R${info['valor']:,.2f}")
            print("Documentos no Banco de Dados:")
            for doc in docs:
                print(f"  - _id: {doc['_id']} (Filial: {doc['filial_codigo']}) | Carregado em: {doc['data_carga']}")

    client.close()

if __name__ == '__main__':
    encontrar_duplicatas_logicas()