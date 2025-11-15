

import os
import glob

import pandas as pd
import os

import chardet

diretorio_atual = os.getcwd()


def localizar_csv(diretorio="."):
    """
    Localiza o primeiro arquivo CSV no diretório especificado.
    """
    arquivos_csv = glob.glob(os.path.join(diretorio, "*.csv"))
    if not arquivos_csv:
        return None
    return arquivos_csv[0]  # Retorna o primeiro arquivo encontrado


def detectar_codificacao(arquivo):
    """
    Detecta a codificação de um arquivo.
    """
    with open(arquivo, 'rb') as f:
        resultado = chardet.detect(f.read(10000))  # Lê os primeiros 10.000 bytes
        return resultado['encoding']

def old_transformar_csv(input_file, output_file):
    """
    Processa o arquivo CSV:
    - Gera uma linha por combinação única de "Código, Nome, Função, Admissão, Situação, Salário".
    - Transforma "Código + Lançamento" em colunas transpostas com os valores da coluna "Valor".
    """
    # Verificar se o arquivo existe
    if not os.path.isfile(input_file):
        print(f"Arquivo não encontrado: {input_file}")
        print("Arquivo precisa ter o nome: planilha_original.csv ")
        input("Pressione Enter para fechar a aplicação...")
        return

    try:

        # Detectar a codificação do arquivo
        encoding_detectado = detectar_codificacao(input_file)
        print(f"Codificação detectada: {encoding_detectado}")


        # Ler o arquivo CSV
        df = pd.read_csv(input_file, sep=';', encoding=encoding_detectado)
        print(f"Colunas existente no arquivo:: {df.columns} \n")

        # Colunas que definem linhas únicas
        colunas_base = ["Código_pessoa", "Nome", "Função", "Admissão", "Situação", "Salário_base"]

        # Colunas que serão usadas para transpor
        colunas_transpostas = ["Código", "Lançamento", "Valor"]

        # Todas as colunas esperadas
        colunas_esperadas = colunas_base + colunas_transpostas

        # Validar se todas as colunas estão presentes no arquivo
        colunas_ausentes = [col for col in colunas_esperadas if col not in df.columns]
        if colunas_ausentes:
            raise KeyError(f"As seguintes colunas estão ausentes no arquivo: {colunas_ausentes}")

        # Agrupar os dados por colunas base
        agrupado = df.groupby(colunas_base)

        # Lista para armazenar as linhas finais
        linhas_finais = []

        for keys, group in agrupado:
            # Criar um dicionário para armazenar os valores agrupados
            linha = {col: keys[i] for i, col in enumerate(colunas_base)}

            # Adicionar colunas transpostas
            for _, row in group[colunas_transpostas].iterrows():
                nome_coluna = f"{row['Lançamento']}"
                linha[nome_coluna] = row["Valor"]

            # Adicionar a linha processada à lista final
            linhas_finais.append(linha)

        # Converter as linhas finais em um DataFrame
        df_final = pd.DataFrame(linhas_finais)

        # Salvar o resultado em um novo arquivo CSV
        df_final.to_csv(output_file, index=False, sep=';', encoding='ISO-8859-1')
        print(f"Arquivo transformado salvo como: {output_file}")


    except KeyError as e:
        print(f"Erro de validação: {e}")
        input("Pressione Enter para fechar a aplicação...")

    except Exception as e:
        print(f"Erro ao processar o arquivo CSV: {e}")
        input("Pressione Enter para fechar a aplicação...")

def transformar_excel(diretorio_atual, output_file):
    """
    Processa o arquivo CSV:
    - Gera uma linha por combinação única de "Código, Nome, Função, Admissão, Situação, Salário".
    - Transforma "Código + Lançamento" em colunas transpostas com os valores da coluna "Valor".
    """
    # Verificar se o arquivo existe
    # Localizar o arquivo CSV automaticamente
    input_file = localizar_csv(diretorio_atual)
    if not os.path.isfile(input_file):
        print(f"Arquivo não encontrado: {input_file}")
        print("Arquivo precisa ser do tipo csv ")
        input("Pressione Enter para fechar a aplicação...")
        return


    try:

        # Detectar a codificação do arquivo
        encoding_detectado = detectar_codificacao(input_file)
        print(f"Arquivo CSV localizado: {input_file}")
        print(f"Codificação detectada: {encoding_detectado} \n")


        # Ler o arquivo CSV
        df = pd.read_csv(input_file, sep=';', encoding=encoding_detectado)
        print(f"Colunas existente no arquivo:: {df.columns} \n")

        # Colunas que definem linhas únicas
        colunas_base = ["Código_pessoa", "Nome", "Função", "Admissão", "Situação", "Salário_base"]

        # Colunas que serão usadas para transpor
        colunas_transpostas = ["Código", "Lançamento", "Valor"]

        # Todas as colunas esperadas
        colunas_esperadas = colunas_base + colunas_transpostas

        # Validar se todas as colunas estão presentes no arquivo
        colunas_ausentes = [col for col in colunas_esperadas if col not in df.columns]
        if colunas_ausentes:
            raise KeyError(f"As seguintes colunas estão ausentes no arquivo: {colunas_ausentes}")

        # Agrupar os dados por colunas base
        agrupado = df.groupby(colunas_base)

        # Lista para armazenar as linhas finais
        linhas_finais = []

        for keys, group in agrupado:
            # Criar um dicionário para armazenar os valores agrupados
            linha = {col: keys[i] for i, col in enumerate(colunas_base)}

            # Adicionar colunas transpostas
            for _, row in group[colunas_transpostas].iterrows():
                nome_coluna = f"{row['Lançamento']}"
                linha[nome_coluna] = row["Valor"]

            # Adicionar a linha processada à lista final
            linhas_finais.append(linha)

        # Converter as linhas finais em um DataFrame
        df_final = pd.DataFrame(linhas_finais)

        # Salvar o resultado diretamente em um arquivo Excel
        df_final.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Arquivo Excel transformado salvo como: {output_file}")


    except KeyError as e:
        print(f"Erro de validação: {e}")
        input("Pressione Enter para fechar a aplicação...")

    except Exception as e:
        print(f"Erro ao processar o arquivo excel: {e}")
        input("Pressione Enter para fechar a aplicação...")


# Exemplo de uso
output_file = fr"{diretorio_atual}\planilha_transformada.xlsx"  # Arquivo de saída
#transformar_csv(input_file, output_file)
transformar_excel(diretorio_atual, output_file)




