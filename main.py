import gradio as gr
import pandas as pd
import re
import os
import matplotlib.pyplot as plt
from datetime import timedelta
from fpdf import FPDF
from weasyprint import HTML
from PyPDF2 import PdfMerger
import numpy as np

def parse_duration(duration_str):
    try:
        h, m, s = map(int, duration_str.split(':'))
        return timedelta(hours=h, minutes=m, seconds=s)
    except:
        return timedelta(0)

def format_timedelta(td):
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def normalize_html_to_csv(input_html_path, output_csv_path):
    html_data = pd.read_html(input_html_path)
    data = html_data[0]
    data.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

def normalize_multiple_excel_to_csv(input_directory, output_directory):
    input_excel_paths = [os.path.join(input_directory, f) for f in os.listdir(input_directory) if f.endswith('.xlsx')]
    output_csv_paths = [os.path.join(output_directory, os.path.splitext(f)[0] + '.csv') for f in os.listdir(input_directory) if f.endswith('.xlsx')]
    for input_excel_path, output_csv_path in zip(input_excel_paths, output_csv_paths):
        excel_data = pd.read_excel(input_excel_path)
        unnecessary_columns = [col for col in excel_data.columns if 'Unnamed' in col]
        if unnecessary_columns:
            excel_data = excel_data.drop(columns=unnecessary_columns)
        excel_data.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

def extract_aluno_pattern(nome):
    if isinstance(nome, str):
        match = re.search(r"(\d{8,9}-\w{2})", nome.lower())
        return match.group(1) if match else None
    return None

def match_alunos(tarefas_csv_path, alunos_csv_path, contador_csv_path):
    try:
        tarefas_df = pd.read_csv(tarefas_csv_path, encoding='utf-8-sig')
        alunos_df = pd.read_csv(alunos_csv_path, encoding='utf-8-sig')
    except pd.errors.EmptyDataError:
        print(f"Arquivo {tarefas_csv_path} ou {alunos_csv_path} está vazio. Pulando...")
        return

    tarefas_df.columns = tarefas_df.columns.str.strip()
    alunos_df.columns = alunos_df.columns.str.strip()

    if 'Aluno' not in tarefas_df.columns or 'Nota' not in tarefas_df.columns or 'Duração' not in tarefas_df.columns:
        print(f"Colunas 'Aluno', 'Nota' ou 'Duração' não encontradas no arquivo {tarefas_csv_path}. Pulando este arquivo.")
        return

    try:
        contador_df = pd.read_csv(contador_csv_path, encoding='utf-8-sig')
    except FileNotFoundError:
        contador_df = pd.DataFrame(columns=['Nome do Aluno', 'Tarefas Completadas', 'Acertos Absolutos', 'Total Tempo'])

    if 'Tarefas Completadas' not in contador_df.columns:
        contador_df['Tarefas Completadas'] = 0
    if 'Acertos Absolutos' not in contador_df.columns:
        contador_df['Acertos Absolutos'] = 0
    if 'Total Tempo' not in contador_df.columns:
        contador_df['Total Tempo'] = '00:00:00'

    def generate_aluno_pattern(ra, dig_ra):
        ra_str = str(ra).zfill(9)
        ra_without_first_two_digits = ra_str[2:]
        return f"{ra_str[1]}{ra_without_first_two_digits}{dig_ra}-sp".lower()

    alunos_df['Aluno_Pattern'] = alunos_df.apply(lambda row: generate_aluno_pattern(row['RA'], row['Dig. RA']), axis=1)

    def extract_aluno_pattern(nome):
        if isinstance(nome, str):
            match = re.search(r'\d+.*', nome.lower())
            return match.group(0) if match else None
        return None

    tarefas_df['Aluno_Pattern'] = tarefas_df['Aluno'].apply(extract_aluno_pattern)
    tarefas_df['Duração'] = tarefas_df['Duração'].apply(parse_duration)

    matched_alunos = alunos_df[alunos_df['Aluno_Pattern'].isin(tarefas_df['Aluno_Pattern'])]

    result_df = matched_alunos[['Nome do Aluno']].drop_duplicates()

    for aluno in result_df['Nome do Aluno']:
        aluno_pattern = alunos_df.loc[alunos_df['Nome do Aluno'] == aluno, 'Aluno_Pattern'].values[0]
        aluno_tarefas = tarefas_df[tarefas_df['Aluno_Pattern'] == aluno_pattern]
        nota_total = aluno_tarefas['Nota'].sum()
        tempo_total = aluno_tarefas['Duração'].sum()

        if aluno in contador_df['Nome do Aluno'].values:
            contador_df.loc[contador_df['Nome do Aluno'] == aluno, 'Tarefas Completadas'] += 1
            contador_df.loc[contador_df['Nome do Aluno'] == aluno, 'Acertos Absolutos'] += nota_total
            current_total_tempo = pd.to_timedelta(contador_df.loc[contador_df['Nome do Aluno'] == aluno, 'Total Tempo'].values[0])
            contador_df.loc[contador_df['Nome do Aluno'] == aluno, 'Total Tempo'] = str(current_total_tempo + tempo_total)
        else:
            contador_df = pd.concat([contador_df, pd.DataFrame({'Nome do Aluno': [aluno], 'Tarefas Completadas': [1], 'Acertos Absolutos': [nota_total], 'Total Tempo': [str(tempo_total)]})], ignore_index=True)

    contador_df.to_csv(contador_csv_path, index=False, encoding='utf-8-sig')

    return result_df

def remove_outliers(data, column, threshold=3):
    mean = data[column].mean()
    std = data[column].std()
    return data[(data[column] > mean - threshold * std) & (data[column] < mean + threshold * std)]

def process_all_tarefas_in_directory(directory, alunos_csv_path, contador_csv_path, relatorio_csv_path):
    tarefas_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.csv') and f not in ['alunos_fim.csv', 'contador_tarefas.csv']]

    for i, tarefas_file in enumerate(tarefas_files):
        match_alunos(tarefas_file, alunos_csv_path, contador_csv_path)

    process_relatorios(contador_csv_path, relatorio_csv_path)

def process_relatorios(contador_csv_path, relatorio_csv_path):
    contador_df = pd.read_csv(contador_csv_path, encoding='utf-8-sig')
    contador_df['Média de Acertos'] = ((contador_df['Acertos Absolutos'] / (contador_df['Tarefas Completadas'] * 2)) * 100).round(2).astype(str) + '%'
    contador_df['Total Tempo'] = pd.to_timedelta(contador_df['Total Tempo'])
    contador_df['Tempo Médio por Tarefa'] = contador_df['Total Tempo'] / contador_df['Tarefas Completadas']
    contador_df['Total Tempo'] = contador_df['Total Tempo'].apply(format_timedelta)
    contador_df['Tempo Médio por Tarefa'] = contador_df['Tempo Médio por Tarefa'].apply(format_timedelta)
    contador_df = contador_df.sort_values(by='Nome do Aluno', ascending=True)

    # Remove outliers e calcula o tempo médio por tarefa da turma
    tempo_medio_por_tarefa = pd.to_timedelta(contador_df['Tempo Médio por Tarefa'])
    tempo_medio_por_tarefa = remove_outliers(pd.DataFrame({'Tempo Médio por Tarefa': tempo_medio_por_tarefa}), 'Tempo Médio por Tarefa')
    media_tempo_medio_turma = tempo_medio_por_tarefa['Tempo Médio por Tarefa'].mean()
    media_tempo_medio_turma = format_timedelta(media_tempo_medio_turma)

    # Calcula médias gerais da turma
    media_acertos_turma = (contador_df['Acertos Absolutos'] / (contador_df['Tarefas Completadas'] * 2)).mean() * 100
    media_tarefas_turma = contador_df['Tarefas Completadas'].mean()

    contador_df.to_csv(relatorio_csv_path, index=False, encoding='utf-8-sig')
    return contador_df, media_tempo_medio_turma, media_acertos_turma, media_tarefas_turma

def generate_pdf_report(dataframe, media_tempo_medio_turma, media_acertos_turma, media_tarefas_turma, html_path, output_pdf_path):
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, 'Relatório de Tarefas', 0, 1, 'C')

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

        def add_image(self, image_path):
            self.add_page()
            self.image(image_path, x=10, y=10, w=270)

    pdf = PDF(orientation='L', unit='mm', format='A4')

    # Gerar gráficos e adicionar ao PDF
    def add_bar_labels(bars, labels):
        for bar, label in zip(bars, labels):
            height = bar.get_height()
            plt.annotate(f'{label}', 
                         xy=(bar.get_x() + bar.get_width() / 2, height), 
                         xytext=(0, 3),  # 3 points vertical offset
                         textcoords="offset points",
                         ha='center', va='bottom')

    top_students = dataframe.nlargest(5, 'Acertos Absolutos')
    
    plt.figure(figsize=(10, 6))
    bars = plt.bar(top_students['Nome do Aluno'], top_students['Acertos Absolutos'], color='blue')
    plt.xlabel('Nome do Aluno')
    plt.ylabel('Acertos Absolutos')
    plt.title('Top 5 Alunos - Acertos Absolutos')
    plt.xticks(rotation=45, ha='right')
    add_bar_labels(bars, top_students['Acertos Absolutos'])
    plt.tight_layout()
    graph_path = 'top_5_acertos_absolutos.png'
    plt.savefig(graph_path)
    pdf.add_image(graph_path)

    plt.figure(figsize=(10, 6))
    bars = plt.bar(top_students['Nome do Aluno'], top_students['Média de Acertos'].str.rstrip('%').astype('float'), color='green')
    plt.xlabel('Nome do Aluno')
    plt.ylabel('Percentual de Acertos (%)')
    plt.title('Top 5 Alunos - Percentual de Acertos')
    plt.xticks(rotation=45, ha='right')
    add_bar_labels(bars, top_students['Média de Acertos'].str.rstrip('%').astype('float'))
    plt.tight_layout()
    graph_path = 'top_5_percentual_acertos.png'
    plt.savefig(graph_path)
    pdf.add_image(graph_path)

    plt.figure(figsize=(10, 6))
    bars = plt.bar(top_students['Nome do Aluno'], top_students['Tarefas Completadas'], color='red')
    plt.xlabel('Nome do Aluno')
    plt.ylabel('Tarefas Completadas')
    plt.title('Top 5 Alunos - Tarefas Completadas')
    plt.xticks(rotation=45, ha='right')
    add_bar_labels(bars, top_students['Tarefas Completadas'])
    plt.tight_layout()
    graph_path = 'top_5_tarefas_completadas.png'
    plt.savefig(graph_path)
    pdf.add_image(graph_path)

    # Adiciona gráfico de alunos que passam mais tempo fazendo as tarefas
    dataframe['Total Tempo'] = pd.to_timedelta(dataframe['Total Tempo'])
    top_time_students = dataframe.nlargest(5, 'Total Tempo')
    plt.figure(figsize=(10, 6))
    bars = plt.bar(top_time_students['Nome do Aluno'], top_time_students['Total Tempo'].dt.total_seconds(), color='purple')
    plt.xlabel('Nome do Aluno')
    plt.ylabel('Tempo Total (hh:mm:ss)')
    plt.title('Top 5 Alunos - Tempo Total')
    plt.xticks(rotation=45, ha='right')
    add_bar_labels(bars, top_time_students['Total Tempo'].apply(format_timedelta))
    plt.tight_layout()
    graph_path = 'top_5_tempo_total.png'
    plt.savefig(graph_path)
    pdf.add_image(graph_path)

    # Adiciona gráfico de resumo da turma
    metrics = ['Tempo Médio (hh:mm:ss)', 'Média de Acertos (%)', 'Média de Tarefas']
    values = [media_tempo_medio_turma, media_acertos_turma, media_tarefas_turma]

    # Convertendo os valores para strings apropriadas
    values_str = [media_tempo_medio_turma, f"{media_acertos_turma:.2f}", f"{media_tarefas_turma:.2f}"]

    plt.figure(figsize=(10, 6))
    bars = plt.bar(metrics, values_str, color=['blue', 'green', 'red'])
    plt.xlabel('Métricas')
    plt.ylabel('Valores')
    plt.title('Resumo da Turma')
    plt.xticks(rotation=45, ha='right')
    for bar, value in zip(bars, values_str):
        plt.annotate(f'{value}', 
                     xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()), 
                     xytext=(0, 3),  # 3 points vertical offset
                     textcoords="offset points",
                     ha='center', va='bottom')
    plt.tight_layout()
    graph_path = 'resumo_turma.png'
    plt.savefig(graph_path)
    pdf.add_image(graph_path)

    # Salvar o PDF com os gráficos
    temp_graphics_pdf = 'temp_graphics.pdf'
    pdf.output(temp_graphics_pdf)

    # Estilo personalizado para bordas da tabela
    html_style = """
    <style>
    .dataframe {
        border-collapse: collapse;
        width: 100%;
    }
    .dataframe th, .dataframe td {
        border: 1px solid black;
        padding: 8px;
        text-align: left;
    }
    </style>
    """

    # Ler o conteúdo HTML e adicionar o estilo personalizado
    with open(html_path, 'r', encoding='utf-8-sig') as f:
        html_content = f.read()
    
    html_content = html_style + html_content

    temp_html_path = 'temp_html_with_borders.html'
    with open(temp_html_path, 'w', encoding='utf-8-sig') as f:
        f.write(html_content)

    temp_html_pdf = 'temp_html.pdf'
    HTML(temp_html_path).write_pdf(temp_html_pdf)

    # Combinar os PDFs
    merger = PdfMerger()
    merger.append(temp_html_pdf)
    merger.append(temp_graphics_pdf)
    merger.write(output_pdf_path)
    merger.close()

    # Remover arquivos temporários
    os.remove(temp_graphics_pdf)
    os.remove(temp_html_pdf)
    os.remove(temp_html_path)

def processar_relatorio(html_file, tarefa_files):
    input_directory = "temp_files"  # Diretório temporário para os arquivos
    output_directory = "temp_files"

    os.makedirs(input_directory, exist_ok=True)
    os.makedirs(output_directory, exist_ok=True)

    # Limpa o diretório temporário antes de cada execução (opcional, mas recomendado)
    for filename in os.listdir(input_directory):
        file_path = os.path.join(input_directory, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

    # Salva os arquivos enviados
    html_path = os.path.join(input_directory, "alunos.htm")
    with open(html_path, "wb") as f:
        f.write(html_file)

    for idx, tarefa_file in enumerate(tarefa_files):
        tarefa_path = os.path.join(input_directory, f"tarefa_{idx}.xlsx")
        with open(tarefa_path, "wb") as f:
            f.write(tarefa_file)

    # Normaliza os arquivos
    alunos_csv_path = os.path.join(output_directory, "alunos_fim.csv")
    normalize_html_to_csv(html_path, alunos_csv_path)
    normalize_multiple_excel_to_csv(input_directory, output_directory)

    # Processa os dados e gera o relatório
    contador_csv_path = os.path.join(output_directory, "contador_tarefas.csv")
    relatorio_csv_path = os.path.join(output_directory, "relatorio_final.csv")
    process_all_tarefas_in_directory(output_directory, alunos_csv_path, contador_csv_path, relatorio_csv_path)
    df, media_tempo_medio_turma, media_acertos_turma, media_tarefas_turma = process_relatorios(contador_csv_path, relatorio_csv_path)

    # Salva o relatório em HTML e PDF
    html_output_path = os.path.join(output_directory, "relatorio_final.html")
    df.to_html(html_output_path, index=False, encoding='utf-8-sig')
    
    pdf_output_path = os.path.join(output_directory, "relatorio_final.pdf")
    generate_pdf_report(df, media_tempo_medio_turma, media_acertos_turma, media_tarefas_turma, html_output_path, pdf_output_path)

    return df.to_html(index=False), html_output_path, pdf_output_path

# Tema personalizado
theme = gr.themes.Default(
    primary_hue="blue",  # Cor principal (tons de azul)
    secondary_hue="gray",  # Cor secundária (tons de cinza)
    font=["Arial", "sans-serif"],  # Família de fontes
    font_mono=["Courier New", "monospace"],  # Fonte para código
)

# --- Interface Gradio ---
with gr.Blocks(theme=theme) as interface:
    gr.Markdown("# Processamento de Relatórios de Tarefas")
    with gr.Row():
        with gr.Column():
            gr.Markdown("## Arquivo HTML (alunos.htm)")
            html_file = gr.File(label="Arraste o arquivo .htm aqui", type="binary")
        with gr.Column():
            gr.Markdown("## Arquivos Excel (Relatórios de Tarefas)")
            excel_files = gr.Files(label="Arraste os arquivos .xlsx aqui", type="binary", file_count="multiple")

    generate_btn = gr.Button("Gerar Relatório", variant="primary")  # Destaque no botão
    output_html = gr.HTML()
    download_html_btn = gr.File(label="Download HTML Report")
    download_pdf_btn = gr.File(label="Download PDF Report")

    def wrapper(html_file, excel_files):
        html_content, html_path, pdf_path = processar_relatorio(html_file, excel_files)
        return html_content, html_path, pdf_path

    generate_btn.click(fn=wrapper, inputs=[html_file, excel_files], outputs=[output_html, download_html_btn, download_pdf_btn])

interface.launch()
