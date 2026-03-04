import pandas as pd
import numpy as np
import os

def gerar_base_looker():
    # 1. Configuração de Caminhos (Diretórios Dinâmicos)
    # Define o caminho baseando-se na localização deste script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)  # Sobe um nível para a raiz do projeto
    
    # Caminho do arquivo de entrada (dentro da pasta /data)
    data_dir = os.path.join(project_root, 'data')
    
    if not os.path.isdir(data_dir):
        print(f"Erro: A pasta 'data' não foi encontrada no caminho esperado: {data_dir}")
        return

    # Busca dinâmica: encontra arquivos CSV ou XLSX na pasta data que contenham "Curva Alunos"
    arquivos_compativeis = [f for f in os.listdir(data_dir) if (f.endswith('.csv') or f.endswith('.xlsx')) and 'Curva Alunos' in f]
    
    if not arquivos_compativeis:
        print(f"Erro: Nenhum arquivo compatível (CSV ou XLSX) contendo 'Curva Alunos' foi encontrado em: {data_dir}")
        todos_os_arquivos = os.listdir(data_dir)
        if todos_os_arquivos:
            print("Arquivos que ESTÃO na pasta 'data':")
            for nome_arquivo in todos_os_arquivos:
                print(f"- {nome_arquivo}")
        else:
            print("A pasta 'data' está vazia.")
        return

    file_path = os.path.join(data_dir, arquivos_compativeis[0])
    print(f"Lendo arquivo: {arquivos_compativeis[0]}")

    # Nome da aba que contém os dados de alunos NOVOS
    sheet_name = 'Curva Alunos NOVO'

    if file_path.endswith('.xlsx'):
        # Tenta ler a aba específica 'Curva Alunos NOVO'
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Aba '{sheet_name}' carregada com sucesso.")
        except ValueError:
            print(f"Aviso: Aba '{sheet_name}' não encontrada. Lendo a primeira aba...")
            df = pd.read_excel(file_path)

        # Normaliza nomes das colunas (remove espaços extras)
        df.columns = df.columns.astype(str).str.strip()

        # Se a coluna chave não for encontrada, tenta ler considerando a segunda linha como cabeçalho
        if 'Ciclo 26.1' not in df.columns:
            print("Aviso: Coluna 'Ciclo 26.1' não encontrada no cabeçalho padrão. Tentando ler com header=1...")
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
            except ValueError:
                df = pd.read_excel(file_path, header=1)
            df.columns = df.columns.astype(str).str.strip()
    else:
        df = pd.read_csv(file_path)
        df.columns = df.columns.astype(str).str.strip()

    if 'Ciclo 26.1' not in df.columns:
        print(f"Erro Crítico: A coluna 'Ciclo 26.1' não foi encontrada. Colunas detectadas: {list(df.columns)}")
        return

    # 2. Limpeza e Formatação de Datas
    # TRANSFORMAÇÃO PARA FORMATO LONGO (STACKED)
    # Empilha os dados de 2026 (Ciclo 26.1) e 2025 (Ciclo 25.1) para ter uma coluna Data única com 2024/2025/2026

    # --- Preparando Ciclo 26.1 (Dados Atuais) e Histórico Consolidado ---
    df_clean = pd.DataFrame()
    df_clean['Data'] = pd.to_datetime(df['Ciclo 26.1'])
    df_clean['Real_Acc'] = df['Alunado 2026 Novo\n(acumulado)']
    df_clean['Meta_Acc'] = df['META MÓVEL NOVO']
    
    # Trazendo Histórico (Ano Anterior) como Coluna (Consolidado)
    # Assume alinhamento por linha na planilha original
    df_clean['Real_Acc_Hist'] = df['Alunado 2025 Novo\n(acumulado)']
    df_clean['Ciclo'] = '26.1'
    
    # Remove linhas onde a Data é nula (linhas vazias do Excel)
    df_clean = df_clean.dropna(subset=['Data']).sort_values('Data')

    # --- Preparando Ciclo 24.1 (Dados de 2024 das colunas I, J, K) ---
    # Identifica as colunas do ciclo anterior (I, J, K)
    ciclo_anterior_cols = [col for col in df.columns if col in df.columns[8:11]]  # Colunas I, J, K (índices 8, 9, 10)
    
    if len(ciclo_anterior_cols) >= 3:
        df_clean_2024 = pd.DataFrame()
        df_clean_2024['Data'] = pd.to_datetime(df[df.columns[8]])  # Coluna I - datas de 2024
        df_clean_2024['Real_Acc'] = df[df.columns[9]]  # Coluna J - Realizado acumulado 2024
        df_clean_2024['Meta_Acc'] = df[df.columns[10]]  # Coluna K - Meta 2024
        df_clean_2024['Real_Acc_Hist'] = np.nan  # Sem histórico para 2024
        df_clean_2024['Ciclo'] = '24.1'
        
        # Remove linhas onde a Data é nula
        df_clean_2024 = df_clean_2024.dropna(subset=['Data']).sort_values('Data')
        
        # Concatena 2024 + 2026
        df_clean = pd.concat([df_clean_2024, df_clean], ignore_index=True).sort_values('Data')

    # 3. Tratamento de Valores Nulos (Forward Fill)
    # Garante que a curva acumulada não tenha quebras (feito separadamente por Ciclo)
    cols_ffill = ['Real_Acc', 'Meta_Acc', 'Real_Acc_Hist']
    df_clean[cols_ffill] = df_clean[cols_ffill].ffill()

    # 4. DINÂMICO: Identificar a data mais recente com dados
    data_mais_recente = df_clean[df_clean['Real_Acc'].notnull()]['Data'].max()
    
    if pd.isnull(data_mais_recente):
        print("Nenhum dado encontrado na planilha.")
        return

    # 5. Cálculo das Métricas Diárias (Transformar acumulado em diário para o Looker)
    # Agrupa por Ciclo para o diff não misturar os anos
    df_clean['Real_Diario'] = df_clean['Real_Acc'].diff().fillna(df_clean['Real_Acc'])
    df_clean['Meta_Diaria'] = df_clean['Meta_Acc'].diff().fillna(df_clean['Meta_Acc'])
    df_clean['Real_Diario_Hist'] = df_clean['Real_Acc_Hist'].diff().fillna(df_clean['Real_Acc_Hist'])

    # 6. Criação de Dimensões de Tempo e Status
    df_clean['Mes'] = df_clean['Data'].dt.month_name()
    
    def get_semana_mes(date):
        # Ajuste: Semana de calendário (início na Segunda-feira)
        first_day = date.replace(day=1)
        day = date.day
        weekday_start = first_day.weekday()

        # Fórmula: (Dia + Dia_Semana_Inicio_Mes - 1) // 7 + 1
        # weekday(): 0=Segunda, 6=Domingo
        week_num = (day + weekday_start - 1) // 7 + 1

        # Se o mês começa no Sábado (5) ou Domingo (6), a primeira "semana" é muito curta.
        # Juntamos com a próxima para que a primeira segunda-feira seja S1.
        if weekday_start >= 5:
            week_num -= 1
            if week_num == 0: week_num = 1

        return 'S' + str(week_num)
    
    df_clean['Semana_Mes'] = df_clean['Data'].apply(get_semana_mes)

    # Nova Coluna: Período da Semana (Segunda a Sexta)
    semana_inicio = df_clean['Data'] - pd.to_timedelta(df_clean['Data'].dt.weekday, unit='D')
    semana_fim = semana_inicio + pd.to_timedelta(4, unit='D')
    df_clean['Semana_Periodo'] = semana_inicio.dt.strftime('%d/%m') + ' - ' + semana_fim.dt.strftime('%d/%m')

    df_clean['Ano_Mes'] = df_clean['Data'].dt.to_period('M').astype(str)
    
    # Status de Performance (Verde/Vermelho no Looker)
    df_clean['Performance_Status'] = np.where(
        df_clean['Real_Diario'] >= df_clean['Meta_Diaria'], 
        'Acima/Na Meta', 
        'Abaixo da Meta'
    )

    # 7. Filtro Automático: Exporta apenas até a última data com dados reais
    # Mantém todo o histórico (25.1) e corta o futuro do 26.1
    base_looker = df_clean[df_clean['Data'] <= data_mais_recente].copy()

    # 8. Exportação para a pasta /outputs (formato Excel para conferência/Looker)
    output_dir = os.path.join(project_root, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    output_name = os.path.join(output_dir, 'db_Dash_curva_de_alunos.xlsx')

    base_looker.to_excel(output_name, index=False)
    
    print(f"Sucesso! Base gerada até {data_mais_recente.strftime('%d/%m/%Y')}: {output_name}")
    
    # Validação final para você ver no console
    anos_encontrados = base_looker['Data'].dt.year.unique()
    print(f"ANOS CONFIRMADOS NA BASE: {list(anos_encontrados)}")
    print(f"Total Realizado (Geral): {base_looker['Real_Diario'].sum():.0f}")
    # print(f"GAP Final Apurado: {base_looker['Real_2026_Diario'].sum() - base_looker['Meta_Diaria'].sum():.0f}")

if __name__ == "__main__":
    gerar_base_looker()