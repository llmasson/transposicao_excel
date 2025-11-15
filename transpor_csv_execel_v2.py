
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
    arquivos_csv = glob.glob(os.path.join(diretorio, "original.csv"))
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
        #df = pd.read_csv(input_file, sep=';', encoding=encoding_detectado)
        df = pd.read_csv(input_file, sep=';', encoding=encoding_detectado, low_memory=False)
        df.columns = df.columns.str.replace(r'[\u200b\u200e\u200f]', '', regex=True).str.strip()

        print(f"Colunas existente no arquivo:: {df.columns} \n")

        # Colunas que definem linhas únicas
        colunas_base = ["codi_emp","nome_emp","cgce_emp","tins_emp","cp_tipo_calc","cp_competencia","cp_data_hora","cp_quebra","cp_pagina_ini","cp_label","cp_codi_epr","cp_nome_epr","cp_situacao","cp_cpf","cp_salario","cp_codi_car","cp_nome_car","cp_vinculo","cp_cc","cp_depto","cp_filial","cp_admissao"]

        # Colunas que serão usadas para transpor
        colunas_transpostas = ["cp_nome_eve_p", "cp_eve_val_p", "cp_nome_eve_d", "cp_eve_val_d",  "cp_tot_inf"]

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
            # Mapeamento das colunas com nome dinâmico
            colunas_pares = [
                ("cp_nome_eve_p", "cp_eve_val_p", "Provento"),
                ("cp_nome_eve_d", "cp_eve_val_d", "Débito"),
            ]

            # Colunas fixas que devem ser transpostas diretamente
            colunas_fixas = [
                "cp_tot_inf"
            ]
            for _, row in group.iterrows():
                for nome_col, valor_col, prefixo in colunas_pares:
                    nome = str(row[nome_col]).strip()
                    valor_str = str(row.get(valor_col,'0'))
                    valor_str = valor_str.replace(".", "").replace(",", ".")
                    valor = float(valor_str)
                    if nome and nome.lower() != "nan" and pd.notna(valor):
                        nome_coluna = f"{prefixo} {nome}"
                        linha[nome_coluna] = valor

                for col in colunas_fixas:
                    valor = row[col]
                    if pd.notna(valor):
                        linha[col] = valor

            # Adicionar a linha processada à lista final
            linhas_finais.append(linha)

        # Converter as linhas finais em um DataFrame
        df_final = pd.DataFrame(linhas_finais)

        # Ordenar as colunas: base → Provento → Débito → outras
        colunas = list(df_final.columns)

        # Garante que colunas base estejam na lista (e na ordem correta)
        colunas_base_presentes = [col for col in colunas_base if col in colunas]

        # Colunas com prefixos
        colunas_provento = sorted([col for col in colunas if col.startswith("Provento ")])
        colunas_debito = sorted([col for col in colunas if col.startswith("Débito ")])

        # Outras colunas restantes
        colunas_restantes = [col for col in colunas if
                             col not in colunas_base_presentes + colunas_provento + colunas_debito]

        # Identifica as colunas por prefixo
        colunas_provento = [col for col in df_final.columns if col.startswith("Provento ")]
        colunas_debito = [col for col in df_final.columns if col.startswith("Débito ")]


        # Nova ordem final
        nova_ordem = colunas_base_presentes + colunas_provento + colunas_debito + colunas_restantes
        df_final = df_final[nova_ordem]


        ######

        # === Somatórios por linha ===
        # identifica as colunas dinâmicas
        colunas_provento = [c for c in df_final.columns if c.startswith("Provento ")]
        colunas_debito = [c for c in df_final.columns if c.startswith("Débito ")]

        # garante que são numéricas (se houver NaN/vazio, vira 0)
        for c in colunas_provento + colunas_debito:
            df_final[c] = pd.to_numeric(df_final[c], errors="coerce").fillna(0)

        # cria as colunas de soma (row-wise)
        df_final["Soma Proventos"] = df_final[colunas_provento].sum(axis=1) if colunas_provento else 0
        df_final["Soma Débitos"] = df_final[colunas_debito].sum(axis=1) if colunas_debito else 0

        ####

        # Salvar o resultado diretamente em um arquivo Excel
        df_final.to_excel(output_file, index=False, engine='openpyxl')
        # Também salva como CSV com separador ;
        csv_output_file = output_file.replace(".xlsx", ".csv")
        df_final.to_csv(csv_output_file, index=False, sep=";", encoding="utf-8-sig")
        print(f"Arquivo CSV também salvo como: {csv_output_file}")

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




