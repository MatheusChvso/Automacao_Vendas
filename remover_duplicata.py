from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"

# -----------------------------------------------------------------------------
# MODO DE SEGURANÇA (DRY RUN)
# -----------------------------------------------------------------------------
# True  = MODO DE SIMULAÇÃO. NENHUM DADO SERÁ APAGADO.
#         O script apenas mostrará quais documentos seriam mantidos e quais seriam deletados.
#         (RECOMENDADO PARA A PRIMEIRA EXECUÇÃO)
#
# False = MODO REAL. O script APAGARÁ PERMANENTEMENTE os dados marcados.
#         (Use apenas DEPOIS de verificar a simulação)
#
DRY_RUN = False
# -----------------------------------------------------------------------------


def limpar_duplicatas_definitivo():
    """Encontra e limpa duplicatas lógicas, mantendo o registro mais recente."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING, serverSelectionTimeoutMS=5000)
        client.admin.command('ping')
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        print(f"Conectado com sucesso ao banco '{MONGO_DATABASE}'.")
    except ConnectionFailure as e:
        print(f"Não foi possível conectar ao MongoDB: {e}")
        return

    # Pipeline para encontrar os grupos de duplicatas lógicas (mesmo do diagnóstico avançado)
    pipeline = [
        {"$group": {
            "_id": {
                "numero_pv": "$numero_pv", "parceiro": "$parceiro",
                "emissao": "$emissao", "valor": "$valor_total_pedido"
            },
            "documentos": {"$push": {"_id": "$_id", "data_carga": "$data_carga"}},
            "count": {"$sum": 1}
        }},
        {"$match": {"count": {"$gt": 1}}}
    ]
    
    print("\nProcurando por grupos de pedidos duplicados para limpeza...")
    grupos_duplicados = list(collection.aggregate(pipeline))

    if not grupos_duplicados:
        print(">> Nenhuma duplicata encontrada para limpar. O banco de dados já está correto.")
        client.close()
        return

    print(f"Encontrados {len(grupos_duplicados)} grupos de duplicatas para processar.")
    total_documentos_removidos = 0

    if DRY_RUN:
        print("\n--- EXECUTANDO EM MODO DE SIMULAÇÃO (DRY RUN) ---")
        print("Nenhum dado será apagado. Apenas mostrando o plano de limpeza.")
    else:
        print("\n--- EXECUTANDO EM MODO REAL ---")
        print("OS DADOS MARCADOS PARA DELEÇÃO SERÃO APAGADOS PERMANENTEMENTE.")

    # Itera sobre cada grupo de duplicatas
    for grupo in grupos_duplicados:
        print("-" * 50)
        info_pedido = grupo['_id']
        print(f"Processando Pedido Nº: {info_pedido['numero_pv']} | Parceiro: {info_pedido['parceiro']}")
        
        # Ordena os documentos pela data de carga, do mais recente para o mais antigo
        documentos_ordenados = sorted(grupo['documentos'], key=lambda x: x['data_carga'], reverse=True)
        
        # O primeiro da lista é o que vamos manter
        documento_para_manter = documentos_ordenados[0]
        
        # Todos os outros são para deletar
        documentos_para_deletar = documentos_ordenados[1:]
        
        print(f"  -> MANTER (mais recente): _id = {documento_para_manter['_id']} (carregado em {documento_para_manter['data_carga']})")
        
        if documentos_para_deletar:
            ids_para_deletar = [doc['_id'] for doc in documentos_para_deletar]
            print(f"  -> DELETAR (mais antigos): {len(ids_para_deletar)} documento(s) com os _ids: {ids_para_deletar}")

            if not DRY_RUN:
                try:
                    resultado = collection.delete_many({"_id": {"$in": ids_para_deletar}})
                    print(f"  -> SUCESSO: {resultado.deleted_count} documento(s) removido(s).")
                    total_documentos_removidos += resultado.deleted_count
                except Exception as e:
                    print(f"  -> ERRO ao deletar: {e}")
        
    print("-" * 50)

    if DRY_RUN:
        print("\nSimulação concluída. Para apagar os dados, mude a variável DRY_RUN para False no script e rode novamente.")
    else:
        print(f"\nLimpeza concluída! Total de {total_documentos_removidos} documentos duplicados removidos.")

    client.close()

if __name__ == '__main__':
    limpar_duplicatas_definitivo()