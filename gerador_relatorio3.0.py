import pandas as pd
from pymongo import MongoClient
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from fpdf import FPDF
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import locale

# --- CONFIGURAÇÕES GERAIS E DE ESTILO ---
# Todas as cores agora estão no formato de texto hexadecimal
MONGO_CONNECTION_STRING = "mongodb://localhost:27017/"
MONGO_DATABASE = "vendas_db"
MONGO_COLLECTION = "pedidos"
COR_PRINCIPAL = "#003f5c"      # Azul Escuro para títulos
COR_SECUNDARIA = "#2f4f4f"    # Cinza Escuro para texto
COR_FUNDO_KPI = "#f0f0f0"      # Cinza Claro para o fundo das caixas de KPI
CORES_GRAFICOS = ["#003f5c", "#58508d", "#bc5090", "#ff6361", "#ffa600"]
FILIAIS_ORDEM = ["Solution", "Vale Aço", "Zona da Mata", "Rio de Janeiro"]
A4_LARGURA = 210
A4_ALTURA = 297
MARGEM = 10
LARGURA_UTIL = A4_LARGURA - (2 * MARGEM)

def hex_to_rgb(hex_color):
    """Converte uma cor de texto hexadecimal (ex: '#RRGGBB') para uma tupla (R, G, B)."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 18)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.cell(0, 10, 'Dashboard de Vendas', 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.set_text_color(128)
        self.cell(0, 8, f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", 0, 1, 'C')
        self.line(MARGEM, self.get_y() + 5, A4_LARGURA - MARGEM, self.get_y() + 5)
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    def caixa_kpi(self, x, y, titulo, valor):
        largura_caixa = (LARGURA_UTIL / 3) - 5
        altura_caixa = 25
        self.set_xy(x, y)
        self.set_fill_color(220, 220, 220)
        self.rect(x + 0.5, y + 0.5, largura_caixa, altura_caixa, 'F')
        self.set_fill_color(*hex_to_rgb(COR_FUNDO_KPI))
        self.set_line_width(0.2)
        self.set_draw_color(220, 220, 220)
        self.rect(x, y, largura_caixa, altura_caixa, 'FD')
        self.set_xy(x, y + 3)
        self.set_font('Arial', 'B', 10)
        self.set_text_color(*hex_to_rgb(COR_SECUNDARIA))
        self.cell(largura_caixa, 8, titulo, 0, 1, 'C')
        self.set_x(x)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.cell(largura_caixa, 10, valor, 0, 1, 'C')

    def criar_tabela(self, x, y, titulo_tabela, header, data, col_widths):
        self.set_xy(x, y)
        self.set_font('Arial', 'B', 11)
        self.set_text_color(*hex_to_rgb(COR_PRINCIPAL))
        self.cell(sum(col_widths), 10, titulo_tabela, 0, 1, 'L')
        self.set_x(x)
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(230, 230, 230)
        for i, col_name in enumerate(header):
            self.cell(col_widths[i], 7, col_name, 1, 0, 'C', fill=True)
        self.ln()
        self.set_font('Arial', '', 7)
        self.set_fill_color(255, 255, 255)
        for row in data:
            self.set_x(x)
            self.cell(col_widths[0], 5, str(row[0]), 1, 0, 'C')
            self.cell(col_widths[1], 5, str(row[1]), 1, 0, 'L')
            self.cell(col_widths[2], 5, str(row[2]), 1, 0, 'L')
            self.cell(col_widths[3], 5, str(row[3]), 1, 0, 'R')
            self.ln()
        return self.get_y()

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

def criar_grafico_vendas_filial(df_dados, titulo, nome_arquivo, tamanho='largo'):
    """Cria um gráfico de barras HORIZONTAL com vendas por filial e salva como imagem."""
    
    # Prepara os dados, garantindo que todas as filiais apareçam mesmo com valor zero
    vendas_por_filial = pd.Series(dtype='float64')
    if not df_dados.empty:
        vendas_por_filial = df_dados.groupby('filial_nome')['valor_total_pedido'].sum()
    
    vendas_por_filial = vendas_por_filial.reindex(FILIAIS_ORDEM).fillna(0)
    
    fig_size = (10, 4.5) if tamanho == 'largo' else (5, 4.5)
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=fig_size)
    
    # <<< ALTERAÇÃO AQUI: Invertendo os eixos X e Y para criar um gráfico horizontal >>>
    sns.barplot(x=vendas_por_filial.values, y=vendas_por_filial.index, palette=CORES_GRAFICOS, ax=ax, orient='h')
    
    ax.set_title(titulo, fontsize=14, weight='bold', pad=20)
    ax.set_xlabel('Vendas (R$)', fontsize=10) # <<< Eixo X agora são os valores
    ax.set_ylabel('') # <<< Remove o rótulo do eixo Y para um visual mais limpo
    
    # Lógica para tratar o eixo X (valores) se houver vendas ou não
    if vendas_por_filial.sum() == 0:
        ax.set_xlim(left=0, right=1)
        ax.set_xticks([0])
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: 'R$ 0'))
    else:
        formatter = mticker.FuncFormatter(lambda x, p: f'R$ {x:,.0f}')
        ax.xaxis.set_major_formatter(formatter) # <<< Formatter aplicado ao eixo X
        for container in ax.containers:
            # A função bar_label se adapta automaticamente para barras horizontais
            ax.bar_label(container, fmt=lambda x: f'R${x:,.0f}', fontsize=8, color='black', padding=3)
        ax.set_xlim(left=0, right=ax.get_xlim()[1] * 1.18)

    # <<< ALTERAÇÃO AQUI: Linha de base agora é vertical em x=0 >>>
    ax.axvline(x=0, color='black', linewidth=1.2)
    
    # <<< REMOVIDO: A rotação dos ticks não é mais necessária >>>
    # plt.xticks(rotation=15, ha='right')
    
    plt.tight_layout()
    
    print(f"Salvando gráfico: {nome_arquivo}")
    plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight')
    plt.close()
    return True

def criar_grafico_evolucao_mensal(df, nome_arquivo):
    """Cria um gráfico de linha com a evolução das vendas nos últimos 12 meses."""
    if df.empty:
        print("Aviso: Sem dados para gerar o gráfico de evolução mensal.")
        return False

    hoje = datetime.now()
    fim_periodo = hoje.replace(day=1) - relativedelta(microseconds=1)
    inicio_periodo = fim_periodo.replace(day=1) - relativedelta(months=11)
    df_periodo = df[(df['emissao'] >= inicio_periodo) & (df['emissao'] <= fim_periodo)]

    if df_periodo.empty:
        print("Aviso: Sem dados nos últimos 12 meses para gerar o gráfico de evolução.")
        return False
        
    vendas_mensais = df_periodo.groupby(pd.Grouper(key='emissao', freq='M'))['valor_total_pedido'].sum()
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, 5))
    
    sns.lineplot(x=vendas_mensais.index, y=vendas_mensais.values, marker='o', color=COR_PRINCIPAL, ax=ax)

    for mes, valor in vendas_mensais.items():
        ax.text(mes, valor + (vendas_mensais.max() * 0.02), f'R${valor:,.0f}', ha='center', size=8, color=COR_SECUNDARIA)

    ax.set_title('Evolução Mensal de Vendas (Últimos 12 Meses)', fontsize=14, weight='bold', pad=20)
    ax.set_xlabel('Mês', fontsize=10)
    ax.set_ylabel('Vendas (R$)', fontsize=10)
    formatter = mticker.FuncFormatter(lambda x, p: f'R$ {x:,.0f}')
    ax.yaxis.set_major_formatter(formatter)
    ax.set_ylim( top=vendas_mensais.max() * 1.15, ymin=0)
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.set_xticks(vendas_mensais.index)
    ax.set_xticklabels(vendas_mensais.index.strftime('%b/%y'), rotation=45, ha='right')
    plt.tight_layout()
    print(f"Salvando gráfico: {nome_arquivo}")
    plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight')
    plt.close()
    return True

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
    
    titulo_grafico_ano = f"Acumulado do Ano ({hoje.year})"
    if not df_ano_atual.empty:
        data_inicio_ano = df_ano_atual['emissao'].min().strftime('%d/%m/%Y')
        data_fim_ano = df_ano_atual['emissao'].max().strftime('%d/%m/%Y')
        titulo_grafico_ano = f"Acumulado do Ano (de {data_inicio_ano} a {data_fim_ano})"

    grafico_ano_ok = criar_grafico_vendas_filial(df_ano_atual, titulo_grafico_ano, 'grafico_ano.png', tamanho='largo')
    grafico_mes_atual_ok = criar_grafico_vendas_filial(df_mes_atual, f"Mês Atual ({hoje.strftime('%B')})", 'grafico_mes_atual.png', tamanho='largo')
    grafico_mes_passado_ok = criar_grafico_vendas_filial(df_mes_passado, f"Mês Anterior ({mes_passado_inicio.strftime('%B')})", 'grafico_mes_passado.png', tamanho='largo')
    grafico_evolucao_ok = criar_grafico_evolucao_mensal(df, 'grafico_evolucao.png')

    pdf = PDF('P', 'mm', 'A4')
    pdf.add_page()
    
    kpi_y_pos = pdf.get_y()
    pdf.caixa_kpi(MARGEM, kpi_y_pos, f"Vendas Mês Atual ({hoje.day} dias)", kpi_mes_atual)
    pdf.caixa_kpi(MARGEM + (LARGURA_UTIL / 3), kpi_y_pos, f"Vendas Mês Passado", kpi_mes_passado)
    pdf.caixa_kpi(MARGEM + 2 * (LARGURA_UTIL / 3), kpi_y_pos, "Vendas no Ano", kpi_ano_atual)
    
    y_pos_atual = kpi_y_pos + 35 
    altura_grafico = 62
    espaco_vertical = 8
    if grafico_ano_ok:
        pdf.image('grafico_ano.png', x=MARGEM, y=y_pos_atual, w=LARGURA_UTIL)
        y_pos_atual += altura_grafico + espaco_vertical
    if grafico_mes_atual_ok:
        pdf.image('grafico_mes_atual.png', x=MARGEM, y=y_pos_atual, w=LARGURA_UTIL)
        y_pos_atual += altura_grafico + espaco_vertical
    if grafico_mes_passado_ok:
        pdf.image('grafico_mes_passado.png', x=MARGEM, y=y_pos_atual, w=LARGURA_UTIL)
    
    pdf.add_page()
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, 'Detalhamento - Últimas 10 Vendas por Filial', 0, 1, 'C')
    pdf.ln(5)
    y_inicio_tabelas = pdf.get_y()
    y_coluna_esquerda_atual = y_inicio_tabelas
    y_coluna_direita_atual = y_inicio_tabelas
    header_tabela = ['Emissão', 'Parceiro', 'Vendedor', 'Valor']
    col_widths = [18, 40, 20, 20]
    for filial in FILIAIS_ORDEM[:2]: 
        df_filial = df[df['filial_nome'] == filial].sort_values(by='emissao', ascending=False).head(10)
        if not df_filial.empty:
            dados = [[row['emissao'].strftime('%d/%m/%y'), row['parceiro'], row['vendedor'], f"R${row['valor_total_pedido']:,.0f}"] for _, row in df_filial.iterrows()]
            y_coluna_esquerda_atual = pdf.criar_tabela(MARGEM, y_coluna_esquerda_atual, f"Últimas Vendas: {filial}", header_tabela, dados, col_widths)
    for filial in FILIAIS_ORDEM[2:]:
        df_filial = df[df['filial_nome'] == filial].sort_values(by='emissao', ascending=False).head(10)
        if not df_filial.empty:
            dados = [[row['emissao'].strftime('%d/%m/%y'), row['parceiro'], row['vendedor'], f"R${row['valor_total_pedido']:,.0f}"] for _, row in df_filial.iterrows()]
            y_coluna_direita_atual = pdf.criar_tabela(A4_LARGURA / 2 + 2, y_coluna_direita_atual, f"Últimas Vendas: {filial}", header_tabela, dados, col_widths)
            
    if grafico_evolucao_ok:
        pdf.add_page()
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Performance de Vendas ao Longo do Tempo', 0, 1, 'C')
        pdf.ln(5)
        pdf.image('grafico_evolucao.png', x=MARGEM, y=pdf.get_y(), w=LARGURA_UTIL)

    nome_pdf_final = f"Dashboard_Vendas_{hoje.strftime('%Y-%m')}.pdf"
    print(f"Salvando PDF final: {nome_pdf_final}")
    pdf.output(nome_pdf_final)

    arquivos_graficos = ['grafico_ano.png', 'grafico_mes_atual.png', 'grafico_mes_passado.png', 'grafico_evolucao.png']
    for f in arquivos_graficos:
        if os.path.exists(f): os.remove(f)

if __name__ == '__main__':
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        print("Aviso: Locale 'pt_BR.UTF-8' não encontrado. Nomes dos meses podem ficar em inglês.")
    
    gerar_relatorio()