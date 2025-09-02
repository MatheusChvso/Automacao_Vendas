import pandas as pd
from pymongo import MongoClient
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from fpdf import FPDF
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

# --- CONFIGURAÇÕES GERAIS ---

# 1. Conexão com o MongoDB
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"

# 2. Padrão de Cores para os Gráficos
CORES_GRAFICOS = ["#003f5c", "#58508d", "#bc5090", "#ff6361", "#ffa600"]

# 3. Nomes das Filiais (para garantir a ordem correta nos gráficos)
FILIAIS_ORDEM = ["Solution", "Vale Aço", "Zona da Mata", "Rio de Janeiro"]


def buscar_dados_mongodb():
    """Busca todos os pedidos do MongoDB e retorna como um DataFrame do Pandas."""
    print("Buscando dados do MongoDB...")
    try:
        client = MongoClient(MONGO_CONNECTION_STRING)
        db = client[MONGO_DATABASE]
        collection = db[MONGO_COLLECTION]
        
        dados = list(collection.find({}))
        
        if not dados:
            print("Aviso: Nenhum dado encontrado no banco de dados.")
            return pd.DataFrame()
        
        df = pd.DataFrame(dados)
        df['emissao'] = pd.to_datetime(df['emissao'])
        print(f"{len(df)} registros encontrados.")
        return df
    except Exception as e:
        print(f"Erro ao buscar dados do MongoDB: {e}")
        return pd.DataFrame()

# <<< FUNÇÃO DE GRÁFICO ATUALIZADA >>>
def criar_grafico_vendas_filial(df_dados, titulo, nome_arquivo):
    """Cria um gráfico de barras vertical com vendas por filial e salva como imagem."""
    if df_dados.empty:
        print(f"Aviso: Sem dados para gerar o gráfico '{titulo}'.")
        return False
        
    vendas_por_filial = df_dados.groupby('filial_nome')['valor_total_pedido'].sum().reindex(FILIAIS_ORDEM).fillna(0)
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(8, 4.5)) # Aumentei um pouco a altura para o rótulo caber melhor

    barplot = sns.barplot(x=vendas_por_filial.index, y=vendas_por_filial.values, palette=CORES_GRAFICOS, ax=ax)
    
    # Título já é centralizado por padrão
    ax.set_title(titulo, fontsize=14, weight='bold', pad=20) # 'pad' adiciona espaço para o título
    ax.set_xlabel('Filial', fontsize=10)
    ax.set_ylabel('Valor Total de Vendas (R$)', fontsize=10)
    
    formatter = mticker.FuncFormatter(lambda x, p: f'R$ {x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    ax.yaxis.set_major_formatter(formatter)

    # <<< ALTERAÇÃO AQUI: Adiciona os rótulos de valor em cima das barras >>>
    for container in ax.containers:
        ax.bar_label(
            container,
            fmt=lambda x: f'R$ {x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'),
            fontsize=9,
            color='black',
            padding=3 # Espaço entre a barra e o texto
        )
    # <<< FIM DA ALTERAÇÃO >>>

    # Ajusta o limite superior do eixo Y para dar espaço para os rótulos
    ax.set_ylim(top=ax.get_ylim()[1] * 1.15)
    
    plt.xticks(rotation=0, ha='center')
    plt.tight_layout()
    
    print(f"Salvando gráfico: {nome_arquivo}")
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()
    return True

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Relatório de Vendas Mensal', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
        
    # <<< FUNÇÃO DE TABELA ATUALIZADA >>>
    def criar_tabela(self, titulo_tabela, header, data):
        self.set_font('Arial', 'B', 12)
        # <<< ALTERAÇÃO AQUI: Título da tabela centralizado >>>
        self.cell(0, 10, titulo_tabela, 0, 1, 'C')
        self.ln(2)

        self.set_font('Arial', 'B', 9)
        col_widths = [25, 85, 30, 30]
        # Calcula o ponto de início X para centralizar a tabela
        table_width = sum(col_widths)
        start_x = self.w / 2 - table_width / 2
        self.set_x(start_x)

        for i, col_name in enumerate(header):
            self.cell(col_widths[i], 7, col_name, 1, 0, 'C')
        self.ln()
        
        self.set_font('Arial', '', 8)
        for row in data:
            self.set_x(start_x) # Reposiciona para cada linha
            # <<< ALTERAÇÃO AQUI: Centraliza o conteúdo das células de dados >>>
            self.cell(col_widths[0], 6, str(row[0]), 1, 0, 'C') # Emissão
            self.cell(col_widths[1], 6, str(row[1]), 1, 0, 'L') # Parceiro (mantido à esquerda para legibilidade)
            self.cell(col_widths[2], 6, str(row[2]), 1, 0, 'C') # Vendedor
            self.cell(col_widths[3], 6, str(row[3]), 1, 0, 'R') # Valor Total (alinhado à direita)
            # <<< FIM DA ALTERAÇÃO >>>
            self.ln()
        self.ln(10)

def gerar_relatorio():
    df = buscar_dados_mongodb()

    if df.empty:
        print("Não foi possível gerar o relatório pois não há dados.")
        return

    hoje = datetime.now()
    mes_atual_inicio = hoje.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    ano_atual_inicio = hoje.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
    mes_passado_fim = mes_atual_inicio - relativedelta(microseconds=1)
    mes_passado_inicio = mes_passado_fim.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
    df_mes_atual = df[df['emissao'] >= mes_atual_inicio]
    df_mes_passado = df[(df['emissao'] >= mes_passado_inicio) & (df['emissao'] <= mes_passado_fim)]
    df_ano_atual = df[df['emissao'] >= ano_atual_inicio]
    
    vendas_mes_atual = df_mes_atual['valor_total_pedido'].sum() if not df_mes_atual.empty else 0
    vendas_mes_passado = df_mes_passado['valor_total_pedido'].sum() if not df_mes_passado.empty else 0
    vendas_ano_atual = df_ano_atual['valor_total_pedido'].sum() if not df_ano_atual.empty else 0
    
    kpi_mes_atual = f"R$ {vendas_mes_atual:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    kpi_mes_passado = f"R$ {vendas_mes_passado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    kpi_ano_atual = f"R$ {vendas_ano_atual:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    grafico1_ok = criar_grafico_vendas_filial(df_mes_atual, f"Vendas por Filial - {hoje.strftime('%B/%Y')}", 'grafico_mes_atual.png')
    grafico2_ok = criar_grafico_vendas_filial(df_ano_atual, f"Vendas por Filial - Acumulado {hoje.year}", 'grafico_ano.png')
    grafico3_ok = criar_grafico_vendas_filial(df_mes_passado, f"Vendas por Filial - {mes_passado_inicio.strftime('%B/%Y')}", 'grafico_mes_passado.png')

    pdf = PDF('P', 'mm', 'A4')
    pdf.add_page()
    
    pdf.set_font('Arial', 'B', 12)
    # <<< ALTERAÇÃO AQUI: Título da seção de KPIs centralizado >>>
    pdf.cell(0, 10, 'Indicadores Chave de Performance (KPIs)', 0, 1, 'C')
    pdf.set_font('Arial', '', 11)
    # O texto dos KPIs em si fica melhor alinhado à esquerda para legibilidade da lista
    pdf.multi_cell(0, 8, 
        f"- Vendas no mês atual ({hoje.strftime('%B/%Y')}, {hoje.day} dias): {kpi_mes_atual}\n"
        f"- Vendas no mês passado ({mes_passado_inicio.strftime('%B/%Y')}): {kpi_mes_passado}\n"
        f"- Vendas acumuladas no ano ({hoje.year}): {kpi_ano_atual}"
    )
    pdf.ln(10)
    
    # Inserção dos gráficos centralizados
    # A lógica x=10, w=190 já centraliza a imagem na área útil da página A4 (210mm de largura com 10mm de margem de cada lado)
    if grafico1_ok:
        pdf.image('grafico_mes_atual.png', x=10, w=190)
    if grafico2_ok:
        pdf.image('grafico_ano.png', x=10, w=190)
    if grafico3_ok:
        pdf.image('grafico_mes_passado.png', x=10, w=190)

    pdf.add_page()
    pdf.set_font('Arial', 'B', 14)
    # <<< ALTERAÇÃO AQUI: Título da página 2 centralizado >>>
    pdf.cell(0, 10, 'Detalhamento - Últimas 10 Vendas por Filial', 0, 1, 'C')
    pdf.ln(5)

    header_tabela = ['Emissão', 'Parceiro', 'Vendedor', 'Valor Total']
    for filial in FILIAIS_ORDEM:
        df_filial = df[df['filial_nome'] == filial].copy()
        df_filial.sort_values(by='emissao', ascending=False, inplace=True)
        ultimas_10_vendas = df_filial.head(10)
        
        if ultimas_10_vendas.empty:
            continue

        dados_tabela = []
        for _, row in ultimas_10_vendas.iterrows():
            dados_tabela.append([
                row['emissao'].strftime('%d/%m/%Y'),
                row['parceiro'],
                row['vendedor'],
                f"R$ {row['valor_total_pedido']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            ])
        
        # A função criar_tabela agora cuida da centralização da tabela e do título
        pdf.criar_tabela(f"Filial: {filial}", header_tabela, dados_tabela)

    nome_pdf_final = f"Relatorio_Vendas_{hoje.strftime('%Y-%m')}.pdf"
    print(f"Salvando PDF final: {nome_pdf_final}")
    pdf.output(nome_pdf_final)

    for f in ['grafico_mes_atual.png', 'grafico_ano.png', 'grafico_mes_passado.png']:
        if os.path.exists(f):
            os.remove(f)

if __name__ == '__main__':
    import locale
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        print("Aviso: Locale 'pt_BR.UTF-8' não encontrado. Nomes dos meses podem ficar em inglês.")
    
    gerar_relatorio()