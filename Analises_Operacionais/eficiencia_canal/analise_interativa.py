
import pandas as pd
import sys

def analisar_leads(caminho_arquivo):
    """
    Carrega um arquivo CSV e permite que o usuário faça perguntas básicas sobre os dados.
    """
    try:
        # Tenta carregar o CSV com diferentes encodings
        try:
            df = pd.read_csv(caminho_arquivo, low_memory=False)
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(caminho_arquivo, low_memory=False, encoding='latin1')
            except UnicodeDecodeError:
                df = pd.read_csv(caminho_arquivo, low_memory=False, encoding='iso-8859-1')

        print("Arquivo CSV carregado com sucesso!")
        print(f"Colunas disponíveis: {', '.join(df.columns)}")
        print("
Digite 'sair' para terminar.")
        print("Exemplos de perguntas:")
        print("  - contar leads onde a coluna 'Nome da Unidade' é 'NOME_DA_UNIDADE'")
        print("  - mostrar leads onde a coluna 'Status do Lead' é 'Aberto'")
        print("  - valores unicos da coluna 'Ciclo de Vida'")

        while True:
            pergunta = input("
Faça sua pergunta: ").lower().strip()

            if pergunta == 'sair':
                print("Encerrando a análise.")
                break

            try:
                if pergunta.startswith("contar leads onde a coluna"):
                    partes = pergunta.split("'")
                    coluna = partes[1]
                    valor = partes[3]
                    contagem = df[df[coluna].str.contains(valor, case=False, na=False)].shape[0]
                    print(f"Resultado: Existem {contagem} leads onde a coluna '{coluna}' contém '{valor}'.")

                elif pergunta.startswith("mostrar leads onde a coluna"):
                    partes = pergunta.split("'")
                    coluna = partes[1]
                    valor = partes[3]
                    resultado = df[df[coluna].str.contains(valor, case=False, na=False)]
                    if not resultado.empty:
                        print(f"Mostrando leads onde a coluna '{coluna}' contém '{valor}':")
                        print(resultado.head())
                    else:
                        print("Nenhum lead encontrado com esse critério.")

                elif pergunta.startswith("valores unicos da coluna"):
                    partes = pergunta.split("'")
                    coluna = partes[1]
                    valores_unicos = df[coluna].unique()
                    print(f"Valores únicos na coluna '{coluna}':")
                    for val in valores_unicos:
                        print(f"- {val}")

                else:
                    print("Pergunta não compreendida. Tente usar um dos formatos de exemplo.")

            except Exception as e:
                print(f"Ocorreu um erro ao processar sua pergunta: {e}")
                print("Verifique se o nome da coluna e os valores estão corretos.")

    except FileNotFoundError:
        print(f"Erro: O arquivo não foi encontrado em '{caminho_arquivo}'")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    # Adicione o caminho completo para o seu arquivo CSV aqui
    caminho_do_arquivo_csv = r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv"
    analisar_leads(caminho_do_arquivo_csv)
