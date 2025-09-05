from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# --- CONFIGURAÇÕES ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"

# MODO DE SEGURANÇA:
# True  = Apenas simula e mostra o que SERIA feito. NENHUM DADO SERÁ ALTERADO.
# False = Executa a migração permanentemente. Use apenas após verificar a simulação.
DRY_RUN = False
# --------------------

def migrar_filiais():
    """Encontra registros de SS e SZM e os funde em JF."""
    try:
        client = MongoClient(MONGO_CONNECTION_STRING, serverSelectionTimeoutMS=5000)
        client.admin.command('ping')
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        print("Conectado com sucesso ao MongoDB.")
    except ConnectionFailure as e:
        print(f"Não foi possível conectar ao MongoDB: {e}")
        return

    filiais_para_migrar = ['SS', 'SZM']
    query = {'filial_codigo': {'$in': filiais_para_migrar}}
    
    documentos_para_migrar = list(collection.find(query))
    
    if not documentos_para_migrar:
        print(">> Nenhum documento das filiais 'SS' ou 'SZM' encontrado para migrar.")
        return

    print(f"Encontrados {len(documentos_para_migrar)} documentos de 'SS' e 'SZM' para fundir em 'JF'.")

    if DRY_RUN:
        print("\n--- EXECUTANDO EM MODO DE SIMULAÇÃO (DRY RUN) ---")
    else:
        print("\n--- EXECUTANDO EM MODO REAL ---")

    for doc in documentos_para_migrar:
        id_antigo = doc['_id']
        id_novo = f"{doc['numero_pv']}_JF"
        
        print(f"Migrando doc {id_antigo} -> {id_novo}")

        if not DRY_RUN:
            try:
                # Remove o _id antigo para que o MongoDB gere um novo se necessário
                doc_original_id = doc.pop('_id')
                
                # Define os novos valores da filial
                doc['filial_codigo'] = 'JF'
                doc['filial_nome'] = 'Juiz de Fora'
                
                # Tenta atualizar (upsert) o documento com o novo _id
                # Se um doc com id_novo já existe, ele será substituído (caso de merge)
                # Se não existe, será criado.
                collection.replace_one(
                    {'_id': id_novo},
                    doc,
                    upsert=True
                )
                
                # Se o id_antigo for diferente do id_novo, apaga o original
                if doc_original_id != id_novo:
                    collection.delete_one({'_id': doc_original_id})
                
                print(f"  -> SUCESSO: Documento {doc_original_id} migrado para {id_novo}.")

            except Exception as e:
                print(f"  -> ERRO ao migrar {id_antigo}: {e}")

    print("\n--- Migração Concluída ---")
    if DRY_RUN:
        print("Simulação finalizada. Para executar de verdade, mude DRY_RUN para False.")

if __name__ == '__main__':
    migrar_filiais()
