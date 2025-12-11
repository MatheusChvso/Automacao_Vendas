import pandas as pd
from pymongo import MongoClient
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from fpdf import FPDF
from datetime import datetime
import os
import locale

# --- CONFIGURAÇÕES VISUAIS & BANCO ---
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"

# Paleta de Cores
COR_PRINCIPAL = "#003f5c"   
COR_SECUNDARIA = "#ffa600"  
COR_TERCIARIA = "#bc5090"   
COR_CINZA = "#58508d"
PALETA_FILIAIS = [COR_PRINCIPAL, COR_SECUNDARIA, COR_TERCIARIA, COR_CINZA]
# Paleta para vendedores (cores distintas)
PALETA_VENDEDORES = ["#003f5c", "#ffa600", "#ff6361", "#bc5090", "#58508d", "#488f31", "#de425b"]

# Configurações PDF (A4)
A4_LARGURA = 210
A4_ALTURA = 297
MARGEM = 10
LARGURA_UTIL = A4_LARGURA - (2 * MARGEM)

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

class PDF(FPDF):
    def __init__(self, ano_relatorio):
        super().__init__('P', 'mm', 'A4')
        self.ano_relatorio = ano_relatorio

    def header(self):
        self.set_font('Arial', 'B', 20)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.cell(0, 10, f'Relatório de Fechamento Anual - {self.ano_relatorio}', 0, 1, 'C')
        
        self.set_font('Arial', 'I', 9)
        self.set_text_color(100)
        self.cell(0, 6, f"Consolidado de Performance de Vendas | Gerado em: {datetime.now().strftime('%d/%m/%Y')}", 0, 1, 'C')
        
        self.set_draw_color(*hex_to_rgb(COR_PRINCIPAL))
        self.set_line_width(0.5)
        self.line(MARGEM, self.get_y() + 2, A4_LARGURA - MARGEM, self.get_y() + 2)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(150)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    def titulo_secao(self, texto):
        self.ln(2)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.set_fill_color(240, 240, 240)
        self.cell(0, 10, f"  {texto}", 0, 1, 'L', fill=True)
        self.ln(3)

    def subtitulo_secao(self, texto):
        self.ln(1)
        self.set_font('Arial', 'B', 11)
        self.set_text_color(80, 80, 80)
        self.cell(0, 8, texto, 0, 1, 'L')

    def criar_kpi_card(self, x, y, w, titulo, valor, cor_texto=COR_PRINCIPAL):
        self.set_xy(x, y)
        self.set_fill_color(252, 252, 252)
        self.set_draw_color(220, 220, 220)
        self.rect(x, y, w, 25, 'FD')
        
        self.set_xy(x, y + 4)
        self.set_font('Arial', '', 9)
        self.set_text_color(80)
        self.cell(w, 5, titulo, 0, 1, 'C')
        
        self.set_x(x)
        self.set_font('Arial', 'B', 13)
        self.set_text_color(*hex_to_rgb(cor_texto))
        self.cell(w, 8, valor, 0, 1, 'C')

def buscar_dados(ano):
    print(f"--- Buscando dados consolidados de {ano} ---")
    client = MongoClient(MONGO_CONNECTION_STRING)
    db = client[MONGO_DATABASE]
    collection = db[MONGO_COLLECTION]
    
    inicio = datetime(ano, 1, 1)
    fim = datetime(ano, 12, 31, 23, 59, 59)
    
    dados = list(collection.find({"emissao": {"$gte": inicio, "$lte": fim}}))
    if not dados: return pd.DataFrame()
    
    df = pd.DataFrame(dados)
    df['emissao'] = pd.to_datetime(df['emissao'])
    df['mes_ano'] = df['emissao'].dt.strftime('%m') 
    return df

def formatar_moeda(valor):
    if valor >= 1_000_000:
        return f'R${valor/1_000_000:.1f}M'
    elif valor >= 1_000:
        return f'R${valor/1_000:.0f}k'
    return f'R${valor:.0f}'

def plotar_evolucao_detalhada(df, nome_arquivo):
    plt.figure(figsize=(12, 6))
    sns.set_style("whitegrid")
    
    vendas_mes = df.groupby(df['emissao'].dt.to_period('M'))['valor_total_pedido'].sum()
    vendas_mes.index = vendas_mes.index.to_timestamp()
    
    ax = sns.lineplot(x=vendas_mes.index, y=vendas_mes.values, marker='o', 
                      markersize=10, linewidth=2.5, color=COR_PRINCIPAL)
    
    max_y = vendas_mes.max()
    for x, y in zip(vendas_mes.index, vendas_mes.values):
        ax.annotate(formatar_moeda(y), (x, y), textcoords="offset points", xytext=(0, 12), 
                    ha='center', fontsize=10, fontweight='bold', color=COR_PRINCIPAL)

    ax.set_title(f'Evolução de Vendas Global ({df["emissao"].dt.year.iloc[0]})', fontsize=14, weight='bold', pad=20)
    ax.set_ylim(0, max_y * 1.15)
    ax.set_ylabel("Vendas", fontsize=10)
    ax.set_xlabel("Mês", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: formatar_moeda(x)))
    plt.xticks(vendas_mes.index, [d.strftime('%b') for d in vendas_mes.index], rotation=0)
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_evolucao_filiais_comparativa(df, nome_arquivo):
    plt.figure(figsize=(12, 7))
    sns.set_style("whitegrid")
    
    df['mes'] = df['emissao'].dt.to_period('M').dt.to_timestamp()
    vendas_pivot = df.groupby(['mes', 'filial_nome'])['valor_total_pedido'].sum().reset_index()
    
    filiais_unicas = sorted(vendas_pivot['filial_nome'].unique())
    cores_map = {filial: PALETA_FILIAIS[i % len(PALETA_FILIAIS)] for i, filial in enumerate(filiais_unicas)}
    
    ax = sns.lineplot(data=vendas_pivot, x='mes', y='valor_total_pedido', hue='filial_nome', 
                      style='filial_nome', markers=True, dashes=False, palette=cores_map, linewidth=2)
    
    for filial in filiais_unicas:
        dados_filial = vendas_pivot[vendas_pivot['filial_nome'] == filial]
        cor_filial = cores_map[filial]
        for idx, row in dados_filial.iterrows():
            ax.annotate(formatar_moeda(row['valor_total_pedido']), (row['mes'], row['valor_total_pedido']),
                        textcoords="offset points", xytext=(0, 8), ha='center', fontsize=8, 
                        fontweight='bold', color=cor_filial, alpha=1.0)

    ax.set_title('Comparativo de Evolução: Filial vs Filial', fontsize=14, weight='bold', pad=20)
    ax.set_ylabel("Vendas")
    ax.set_xlabel("")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: formatar_moeda(x)))
    plt.xticks(vendas_pivot['mes'].unique(), [pd.to_datetime(d).strftime('%b') for d in vendas_pivot['mes'].unique()])
    plt.legend(title='Filial', loc='upper left', bbox_to_anchor=(1, 1))
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_evolucao_vendedores_fatiado(df, nome_arquivo, rank_inicio, rank_fim, titulo_custom=None):
    """
    Plota evolução de um slice específico de vendedores (ex: 1º ao 3º).
    rank_inicio e rank_fim são índices base 0 (0 = 1º lugar).
    """
    plt.figure(figsize=(12, 6)) # Altura reduzida para caber 2 na página
    sns.set_style("whitegrid")

    # 1. Identificar TODOS os Top Vendedores para manter consistência de cores se necessário
    # ou apenas pega o ranking geral
    ranking_geral = df.groupby('vendedor')['valor_total_pedido'].sum().sort_values(ascending=False).index.tolist()
    
    # 2. Selecionar apenas o intervalo desejado
    # Proteção para não estourar lista
    if rank_inicio >= len(ranking_geral):
        return False # Nada a plotar
        
    vendedores_selecionados = ranking_geral[rank_inicio : min(rank_fim, len(ranking_geral))]
    
    if not vendedores_selecionados:
        return False

    # 3. Filtrar Dados
    df_slice = df[df['vendedor'].isin(vendedores_selecionados)].copy()
    df_slice['mes'] = df_slice['emissao'].dt.to_period('M').dt.to_timestamp()
    
    # 4. Agrupar
    vendas_pivot = df_slice.groupby(['mes', 'vendedor'])['valor_total_pedido'].sum().reset_index()
    
    # Mapeamento de Cores: Usar índice do ranking geral para garantir cores consistentes ou fixas
    # Mas aqui vamos usar uma paleta local para garantir contraste neste gráfico específico
    cores_map = {vend: PALETA_VENDEDORES[i % len(PALETA_VENDEDORES)] for i, vend in enumerate(vendedores_selecionados)}

    # 5. Plotar
    ax = sns.lineplot(data=vendas_pivot, x='mes', y='valor_total_pedido', hue='vendedor', 
                      marker='o', palette=cores_map, linewidth=2.5)

    # 6. Anotações
    for vendedor in vendedores_selecionados:
        dados_vend = vendas_pivot[vendas_pivot['vendedor'] == vendedor]
        cor_vend = cores_map[vendedor]
        for idx, row in dados_vend.iterrows():
            if row['valor_total_pedido'] > 0:
                ax.annotate(formatar_moeda(row['valor_total_pedido']), 
                           (row['mes'], row['valor_total_pedido']),
                           textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8, 
                           fontweight='bold', color=cor_vend)

    titulo = titulo_custom if titulo_custom else f'Top {rank_inicio+1} a {rank_fim} Vendedores'
    ax.set_title(titulo, fontsize=12, weight='bold', pad=15)
    ax.set_ylabel("Vendas")
    ax.set_xlabel("")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: formatar_moeda(x)))
    
    unique_meses = sorted(vendas_pivot['mes'].unique())
    plt.xticks(unique_meses, [pd.to_datetime(d).strftime('%b') for d in unique_meses])
    
    # Legenda compacta
    plt.legend(title='', loc='upper left', bbox_to_anchor=(1, 1), fontsize='small')
    
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()
    return True

def plotar_vendedores_ranking(df, nome_arquivo):
    plt.figure(figsize=(10, 8))
    vendas_vendedor = df.groupby('vendedor')['valor_total_pedido'].sum().sort_values(ascending=False)
    top_vendas = vendas_vendedor.head(15)
    
    ax = sns.barplot(x=top_vendas.values, y=top_vendas.index, palette="Blues_r")
    ax.set_title('Ranking Geral - Top 15 Vendedores (Acumulado)', fontsize=14, weight='bold')
    ax.set_xlabel("Volume de Vendas (R$)", fontsize=10)
    ax.set_ylabel("")
    
    for container in ax.containers:
        ax.bar_label(container, fmt=lambda x: f' R$ {x:,.0f}', fontsize=9, padding=3)
        
    sns.despine(left=True, bottom=True)
    ax.set_xticklabels([])
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_distribuicao_filiais(df, nome_arquivo):
    plt.figure(figsize=(7, 7))
    total_por_filial = df.groupby('filial_nome')['valor_total_pedido'].sum()
    plt.pie(total_por_filial, labels=None, autopct='%1.1f%%', startangle=90, pctdistance=0.85, colors=PALETA_FILIAIS)
    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    plt.legend(total_por_filial.index, loc="center", bbox_to_anchor=(0.5, 0.5), frameon=False)
    plt.title('Share de Vendas por Filial', weight='bold')
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def gerar_tabela_resumo_mensal(pdf, df):
    pdf.titulo_secao("Matriz de Vendas Mensal")
    df['mes_num'] = df['emissao'].dt.month
    matriz_simples = df.groupby(['mes_num', 'filial_nome'])['valor_total_pedido'].sum().unstack(fill_value=0)
    
    header = ['Mês'] + list(matriz_simples.columns) + ['TOTAL MÊS']
    col_w_base = LARGURA_UTIL / len(header)
    col_widths = [col_w_base] * len(header)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(*hex_to_rgb(COR_PRINCIPAL))
    pdf.set_text_color(255)
    for i, h in enumerate(header):
        pdf.cell(col_widths[i], 8, str(h).upper(), 1, 0, 'C', fill=True)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    pdf.set_text_color(0)
    total_geral_anual = 0
    
    for mes_num in sorted(matriz_simples.index):
        row = matriz_simples.loc[mes_num]
        mes_nome = datetime(2000, mes_num, 1).strftime('%B')
        pdf.set_fill_color(245, 245, 245) if mes_num % 2 == 0 else pdf.set_fill_color(255)
        pdf.cell(col_widths[0], 7, mes_nome.title(), 1, 0, 'C', fill=True)
        total_mes = 0
        for i, filial in enumerate(matriz_simples.columns):
            valor = row[filial]
            total_mes += valor
            pdf.cell(col_widths[i+1], 7, f"R$ {valor:,.2f}", 1, 0, 'R', fill=True)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(col_widths[-1], 7, f"R$ {total_mes:,.2f}", 1, 0, 'R', fill=True)
        pdf.set_font('Arial', '', 8)
        pdf.ln()
        total_geral_anual += total_mes

    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(*hex_to_rgb(COR_SECUNDARIA))
    pdf.cell(col_widths[0], 8, "TOTAL ANUAL", 1, 0, 'C', fill=True)
    for i, filial in enumerate(matriz_simples.columns):
        val_anual = matriz_simples[filial].sum()
        pdf.cell(col_widths[i+1], 8, f"R$ {val_anual:,.2f}", 1, 0, 'R', fill=True)
    pdf.cell(col_widths[-1], 8, f"R$ {total_geral_anual:,.2f}", 1, 0, 'R', fill=True)
    pdf.ln()

def gerar_relatorio_final(ano_alvo):
    # 1. Obter Dados
    df = buscar_dados(ano_alvo)
    if df.empty:
        print(f"ERRO: Nenhum dado encontrado para o ano {ano_alvo}.")
        return

    print("Gerando gráficos...")
    plotar_evolucao_detalhada(df, 'chart_evolucao_geral.png')
    plotar_evolucao_filiais_comparativa(df, 'chart_evolucao_comp.png')
    
    # SPLIT DOS VENDEDORES
    # Gráfico 1: Top 1 ao 3
    plotar_evolucao_vendedores_fatiado(df, 'chart_vend_1_3.png', 0, 3, "Evolução Mensal - Top 3 Vendedores")
    # Gráfico 2: Top 4 ao 5
    plotar_evolucao_vendedores_fatiado(df, 'chart_vend_4_5.png', 3, 5, "Evolução Mensal - Vendedores 4º e 5º")
    
    plotar_vendedores_ranking(df, 'chart_vendedores.png')
    plotar_distribuicao_filiais(df, 'chart_pizza.png')

    # 2. Construir PDF
    print("Montando o PDF...")
    pdf = PDF(ano_alvo)
    
    # --- PÁGINA 1: VISÃO GERAL ---
    pdf.add_page()
    total_vendas = df['valor_total_pedido'].sum()
    total_pedidos = len(df)
    melhor_mes = df.groupby(df['emissao'].dt.strftime('%m/%Y'))['valor_total_pedido'].sum().idxmax()
    
    y_kpi = pdf.get_y()
    espaco = 2
    w_kpi = (LARGURA_UTIL - (espaco * 2)) / 3 
    pdf.criar_kpi_card(MARGEM, y_kpi, w_kpi, "Total de Vendas", f"R$ {total_vendas:,.2f}")
    pdf.criar_kpi_card(MARGEM + w_kpi + espaco, y_kpi, w_kpi, "Total Pedidos", f"{total_pedidos}")
    pdf.criar_kpi_card(MARGEM + (w_kpi + espaco)*2, y_kpi, w_kpi, "Melhor Mês", f"{melhor_mes}")
    pdf.ln(30)
    
    pdf.titulo_secao("Evolução Mensal de Vendas (Geral)")
    pdf.image('chart_evolucao_geral.png', x=MARGEM, w=LARGURA_UTIL, h=90)
    pdf.ln(5)
    
    pdf.titulo_secao("Performance Mensal por Filial")
    pdf.image('chart_evolucao_comp.png', x=MARGEM, w=LARGURA_UTIL, h=95)

    # --- PÁGINA 2: VENDEDORES (Foco total na equipe) ---
    pdf.add_page()
    pdf.titulo_secao("Análise de Performance: Vendedores")
    
    # Subtítulo e Gráfico Top 3
    # pdf.subtitulo_secao("Destaques Principais (Top 3)")
    y_cursor = pdf.get_y()
    pdf.image('chart_vend_1_3.png', x=MARGEM, y=y_cursor, w=LARGURA_UTIL, h=80)
    pdf.set_y(y_cursor + 82)
    
    # Subtítulo e Gráfico Top 4-5
    # pdf.subtitulo_secao("Seguidores (4º e 5º Colocados)")
    y_cursor = pdf.get_y()
    pdf.image('chart_vend_4_5.png', x=MARGEM, y=y_cursor, w=LARGURA_UTIL, h=80)
    pdf.set_y(y_cursor + 85)
    
    # Ranking Geral
    # pdf.subtitulo_secao("Ranking Acumulado")
    pdf.image('chart_vendedores.png', x=MARGEM, y=pdf.get_y(), w=LARGURA_UTIL, h=95)

    # --- PÁGINA 3: FILIAIS & CLIENTES ---
    pdf.add_page()
    pdf.titulo_secao("Detalhamento por Unidade e Clientes")
    
    y_charts = pdf.get_y()
    pdf.image('chart_pizza.png', x=MARGEM, y=y_charts, w=80)
    
    pdf.set_y(y_charts + 85)
    gerar_tabela_resumo_mensal(pdf, df)
    
    pdf.ln(10)
    pdf.titulo_secao("Principais Parceiros (Top 20 Clientes)")
    
    top_clientes = df.groupby('parceiro')['valor_total_pedido'].sum().sort_values(ascending=False).head(20).reset_index()
    pdf.set_fill_color(*hex_to_rgb(COR_PRINCIPAL))
    pdf.set_text_color(255)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(15, 8, "#", 1, 0, 'C', fill=True)
    pdf.cell(120, 8, "PARCEIRO", 1, 0, 'L', fill=True)
    pdf.cell(55, 8, "TOTAL COMPRADO", 1, 1, 'R', fill=True)
    pdf.set_text_color(0)
    pdf.set_font('Arial', '', 9)
    for i, row in top_clientes.iterrows():
        if i % 2 == 0: pdf.set_fill_color(245)
        else: pdf.set_fill_color(255)
        pdf.cell(15, 7, str(i+1), 1, 0, 'C', fill=True)
        pdf.cell(120, 7, row['parceiro'][:50], 1, 0, 'L', fill=True)
        pdf.cell(55, 7, f"R$ {row['valor_total_pedido']:,.2f}", 1, 1, 'R', fill=True)

    # 3. Finalizar
    nome_arquivo_pdf = f"Fechamento_Anual_{ano_alvo}.pdf"
    pdf.output(nome_arquivo_pdf)
    
    # Limpeza
    for f in ['chart_evolucao_geral.png', 'chart_evolucao_comp.png', 'chart_vend_1_3.png', 
              'chart_vend_4_5.png', 'chart_vendedores.png', 'chart_pizza.png']:
        if os.path.exists(f): os.remove(f)

    print(f"\n--- SUCESSO! Relatório salvo como: {nome_arquivo_pdf} ---")

if __name__ == '__main__':
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except:
        print("Aviso: Locale PT-BR não configurado.")
    ano = int(input("Digite o ano para o fechamento (ex: 2024): "))
    gerar_relatorio_final(ano)