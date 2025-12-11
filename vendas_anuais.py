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

# Paleta de Cores (Identidade Visual Mantida)
COR_PRINCIPAL = "#003f5c"   # Azul Escuro
COR_SECUNDARIA = "#ffa600"  # Laranja/Amarelo
COR_TERCIARIA = "#bc5090"   # Roxo
COR_CINZA = "#58508d"
PALETA_FILIAIS = [COR_PRINCIPAL, COR_SECUNDARIA, COR_TERCIARIA, COR_CINZA]

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
        # Cabeçalho limpo e profissional
        self.set_font('Arial', 'B', 20)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.cell(0, 10, f'Relatório de Fechamento Anual - {self.ano_relatorio}', 0, 1, 'C')
        
        self.set_font('Arial', 'I', 9)
        self.set_text_color(100)
        self.cell(0, 6, f"Consolidado de Performance de Vendas | Gerado em: {datetime.now().strftime('%d/%m/%Y')}", 0, 1, 'C')
        
        # Linha divisória elegante
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
        self.ln(5)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.set_fill_color(240, 240, 240) # Fundo cinza claro para destaque
        self.cell(0, 10, f"  {texto}", 0, 1, 'L', fill=True)
        self.ln(3)

    def criar_kpi_card(self, x, y, w, titulo, valor, cor_texto=COR_PRINCIPAL):
        self.set_xy(x, y)
        # Sombra/Borda suave
        self.set_fill_color(252, 252, 252)
        self.set_draw_color(220, 220, 220)
        self.rect(x, y, w, 25, 'FD')
        
        # Título KPI
        self.set_xy(x, y + 4)
        self.set_font('Arial', '', 9)
        self.set_text_color(80)
        self.cell(w, 5, titulo, 0, 1, 'C')
        
        # Valor KPI
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
    # Cria coluna Mês/Ano para agrpamento
    df['mes_ano'] = df['emissao'].dt.strftime('%m') 
    return df

def formatar_moeda(valor):
    if valor >= 1_000_000:
        return f'R${valor/1_000_000:.1f}M'
    elif valor >= 1_000:
        return f'R${valor/1_000:.0f}k'
    return f'R${valor:.0f}'

def plotar_evolucao_detalhada(df, nome_arquivo):
    """Gera gráfico de linha com valores escritos explicitamente em cada ponto."""
    plt.figure(figsize=(12, 6))
    sns.set_style("whitegrid")
    
    # Agrupa por mês
    vendas_mes = df.groupby(df['emissao'].dt.to_period('M'))['valor_total_pedido'].sum()
    vendas_mes.index = vendas_mes.index.to_timestamp()
    
    # Plotagem
    ax = sns.lineplot(x=vendas_mes.index, y=vendas_mes.values, marker='o', 
                      markersize=10, linewidth=2.5, color=COR_PRINCIPAL)
    
    # Anotações de Valor (O que você pediu)
    max_y = vendas_mes.max()
    for x, y in zip(vendas_mes.index, vendas_mes.values):
        ax.annotate(formatar_moeda(y), 
                    (x, y), 
                    textcoords="offset points", 
                    xytext=(0, 12), 
                    ha='center', 
                    fontsize=10, 
                    fontweight='bold', 
                    color=COR_PRINCIPAL)

    # Configurações do Eixo
    ax.set_title(f'Evolução de Faturamento Global ({df["emissao"].dt.year.iloc[0]})', fontsize=14, weight='bold', pad=20)
    ax.set_ylim(0, max_y * 1.15) # Dá espaço para os rótulos não cortarem
    ax.set_ylabel("Faturamento", fontsize=10)
    ax.set_xlabel("Mês", fontsize=10)
    
    # Formata eixo Y e X
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: formatar_moeda(x)))
    plt.xticks(vendas_mes.index, [d.strftime('%b') for d in vendas_mes.index], rotation=0)
    
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_evolucao_filiais_comparativa(df, nome_arquivo):
    """Gráfico multilinhas: Uma linha para cada filial, com valores."""
    plt.figure(figsize=(12, 7))
    sns.set_style("whitegrid")
    
    # Prepara dados: Pivot Table para ter meses nas colunas e filiais nas linhas (ou hue)
    df['mes'] = df['emissao'].dt.to_period('M').dt.to_timestamp()
    vendas_pivot = df.groupby(['mes', 'filial_nome'])['valor_total_pedido'].sum().reset_index()
    
    # Plotagem
    ax = sns.lineplot(data=vendas_pivot, x='mes', y='valor_total_pedido', hue='filial_nome', 
                      style='filial_nome', markers=True, dashes=False, palette=PALETA_FILIAIS, linewidth=2)
    
    # Anotar apenas os pontos de pico ou final? Não, o usuário quer detalhe.
    # Vamos anotar todos, mas com cuidado para não sobrepor demais.
    filiais = vendas_pivot['filial_nome'].unique()
    
    for filial in filiais:
        dados_filial = vendas_pivot[vendas_pivot['filial_nome'] == filial]
        for idx, row in dados_filial.iterrows():
            ax.annotate(formatar_moeda(row['valor_total_pedido']), 
                        (row['mes'], row['valor_total_pedido']),
                        textcoords="offset points",
                        xytext=(0, 8),
                        ha='center',
                        fontsize=8,
                        color='black',
                        alpha=0.7)

    ax.set_title('Comparativo de Evolução: Filial vs Filial', fontsize=14, weight='bold', pad=20)
    ax.set_ylabel("")
    ax.set_xlabel("")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: formatar_moeda(x)))
    plt.xticks(vendas_pivot['mes'].unique(), [pd.to_datetime(d).strftime('%b') for d in vendas_pivot['mes'].unique()])
    plt.legend(title='Filial', loc='upper left', bbox_to_anchor=(1, 1))
    
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_vendedores_ranking(df, nome_arquivo):
    """Barra horizontal dos Top Vendedores."""
    plt.figure(figsize=(10, 8)) # Mais alto para caber os nomes
    
    vendas_vendedor = df.groupby('vendedor')['valor_total_pedido'].sum().sort_values(ascending=False)
    # Pega top 15 para não poluir
    top_vendas = vendas_vendedor.head(15)
    
    ax = sns.barplot(x=top_vendas.values, y=top_vendas.index, palette="Blues_r")
    
    ax.set_title('Top 15 Vendedores - Acumulado Anual', fontsize=14, weight='bold')
    ax.set_xlabel("Volume de Vendas (R$)", fontsize=10)
    ax.set_ylabel("")
    
    # Labels nas barras
    for container in ax.containers:
        ax.bar_label(container, fmt=lambda x: f' R$ {x:,.0f}', fontsize=9, padding=3)
        
    sns.despine(left=True, bottom=True)
    ax.set_xticklabels([]) # Remove eixo X numérico para limpar
    
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def plotar_distribuicao_filiais(df, nome_arquivo):
    """Pizza/Donut para share de filiais."""
    plt.figure(figsize=(7, 7))
    total_por_filial = df.groupby('filial_nome')['valor_total_pedido'].sum()
    
    plt.pie(total_por_filial, labels=None, autopct='%1.1f%%', startangle=90, pctdistance=0.85, colors=PALETA_FILIAIS)
    
    # Círculo central para transformar em Donut Chart (mais moderno)
    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    
    plt.legend(total_por_filial.index, loc="center", bbox_to_anchor=(0.5, 0.5), frameon=False)
    plt.title('Representatividade por Filial (Market Share)', weight='bold')
    plt.tight_layout()
    plt.savefig(nome_arquivo, dpi=300)
    plt.close()

def gerar_tabela_resumo_mensal(pdf, df):
    """Cria uma matriz Mês x Filial no PDF."""
    pdf.titulo_secao("Matriz de Faturamento Mensal")
    
    # Pivota os dados
    df['mes_num'] = df['emissao'].dt.month
    df['mes_nome'] = df['emissao'].dt.strftime('%B')
    matriz = df.groupby(['mes_num', 'mes_nome', 'filial_nome'])['valor_total_pedido'].sum().unstack(fill_value=0)
    
    # Configurações da Tabela
    header = ['Mês'] + list(matriz.columns) + ['TOTAL MÊS']
    col_w_base = LARGURA_UTIL / len(header)
    col_widths = [col_w_base] * len(header)
    
    # Desenha Header
    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(*hex_to_rgb(COR_PRINCIPAL))
    pdf.set_text_color(255)
    for i, h in enumerate(header):
        pdf.cell(col_widths[i], 8, str(h).upper(), 1, 0, 'C', fill=True)
    pdf.ln()
    
    # Desenha Linhas
    pdf.set_font('Arial', '', 8)
    pdf.set_text_color(0)
    
    total_geral_anual = 0
    
    for idx, row in matriz.iterrows():
        # idx é (mes_num, mes_nome) se agrupamos assim, ou apenas mes_num se o index for simples
        # Como o index original do unstack manteve mes_num e mes_nome, precisamos tratar
        # Vamos simplificar o loop iterando pelos indices ordenados
        pass
    
    # Re-agrupando para iterar limpo
    matriz_simples = df.groupby(['mes_num', 'filial_nome'])['valor_total_pedido'].sum().unstack(fill_value=0)
    
    for mes_num in sorted(matriz_simples.index):
        row = matriz_simples.loc[mes_num]
        mes_nome = datetime(2000, mes_num, 1).strftime('%B') # Pega nome do mês
        
        pdf.set_fill_color(245, 245, 245) if mes_num % 2 == 0 else pdf.set_fill_color(255)
        
        # Coluna Mês
        pdf.cell(col_widths[0], 7, mes_nome.title(), 1, 0, 'C', fill=True)
        
        total_mes = 0
        # Colunas Filiais
        for i, filial in enumerate(matriz_simples.columns):
            valor = row[filial]
            total_mes += valor
            pdf.cell(col_widths[i+1], 7, f"R$ {valor:,.2f}", 1, 0, 'R', fill=True)
        
        # Coluna Total
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(col_widths[-1], 7, f"R$ {total_mes:,.2f}", 1, 0, 'R', fill=True)
        pdf.set_font('Arial', '', 8)
        pdf.ln()
        total_geral_anual += total_mes

    # Linha Total Anual
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

    print("Gerando gráficos de alta resolução...")
    plotar_evolucao_detalhada(df, 'chart_evolucao_geral.png')
    plotar_evolucao_filiais_comparativa(df, 'chart_evolucao_comp.png')
    plotar_vendedores_ranking(df, 'chart_vendedores.png')
    plotar_distribuicao_filiais(df, 'chart_pizza.png')

    # 2. Construir PDF
    print("Montando o PDF...")
    pdf = PDF(ano_alvo)
    
    # --- PÁGINA 1: VISÃO GERAL ---
    pdf.add_page()
    
    # KPIs Rápidos
    total_vendas = df['valor_total_pedido'].sum()
    total_pedidos = len(df)
    ticket_medio = total_vendas / total_pedidos if total_pedidos > 0 else 0
    melhor_mes = df.groupby(df['emissao'].dt.strftime('%m/%Y'))['valor_total_pedido'].sum().idxmax()
    
    y_kpi = pdf.get_y()
    w_kpi = LARGURA_UTIL / 4 - 2
    pdf.criar_kpi_card(MARGEM, y_kpi, w_kpi, "Faturamento Anual", f"R$ {total_vendas:,.2f}")
    pdf.criar_kpi_card(MARGEM + w_kpi + 2, y_kpi, w_kpi, "Total Pedidos", f"{total_pedidos}")
    pdf.criar_kpi_card(MARGEM + (w_kpi + 2)*2, y_kpi, w_kpi, "Ticket Médio", f"R$ {ticket_medio:,.2f}")
    pdf.criar_kpi_card(MARGEM + (w_kpi + 2)*3, y_kpi, w_kpi, "Melhor Mês", f"{melhor_mes}")
    
    pdf.ln(30)
    
    # Gráfico Evolução Principal
    pdf.titulo_secao("Evolução Mensal do Faturamento")
    pdf.image('chart_evolucao_geral.png', x=MARGEM, w=LARGURA_UTIL, h=90)
    
    pdf.ln(5)
    
    # Gráfico Comparativo Filiais (Linha)
    pdf.titulo_secao("Performance Mensal por Filial")
    pdf.image('chart_evolucao_comp.png', x=MARGEM, w=LARGURA_UTIL, h=95)

    # --- PÁGINA 2: DETALHE VENDEDORES & FILIAIS ---
    pdf.add_page()
    
    # Divide a página: Esquerda (Pizza), Direita (Vendedores)
    pdf.titulo_secao("Análise de Equipe e Unidades")
    
    y_charts = pdf.get_y()
    
    # Donut Chart Filiais
    pdf.image('chart_pizza.png', x=MARGEM, y=y_charts, w=80)
    
    # Ranking Vendedores (ocupando a direita e descendo mais)
    pdf.image('chart_vendedores.png', x=MARGEM + 85, y=y_charts, w=105, h=110)
    
    # Move cursor para baixo dos gráficos
    pdf.set_y(y_charts + 115)
    
    # Tabela Matriz Mensal
    gerar_tabela_resumo_mensal(pdf, df)
    
    # --- PÁGINA 3: TOP CLIENTES (Pareto) ---
    pdf.add_page()
    pdf.titulo_secao("Principais Parceiros (Top 20 Clientes)")
    
    top_clientes = df.groupby('parceiro')['valor_total_pedido'].sum().sort_values(ascending=False).head(20).reset_index()
    
    # Cabeçalho Tabela Clientes
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
    for f in ['chart_evolucao_geral.png', 'chart_evolucao_comp.png', 'chart_vendedores.png', 'chart_pizza.png']:
        if os.path.exists(f): os.remove(f)

    print(f"\n--- SUCESSO! Relatório salvo como: {nome_arquivo_pdf} ---")

if __name__ == '__main__':
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except:
        print("Aviso: Locale PT-BR não configurado no sistema. Meses podem aparecer em inglês.")
        
    ano = int(input("Digite o ano para o fechamento (ex: 2024): "))
    gerar_relatorio_final(ano)