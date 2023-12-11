import os
import pandas as pd
import glob
import locale


locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

base_dir = "Diretorio Base"
caminho_excel = f"{base_dir}\\cnpjs_com_nomes_status.xlsx"
caminho_excel_copia = f"{base_dir}\\cnpjs_com_nomes_status_copia.xlsx"

df = pd.read_excel(caminho_excel, sheet_name='BASE')
excel_copia = df.copy()


def extrair_ano_mes(coluna):
    meses_numeros = {
        'JANEIRO': 1,
        'FEVEREIRO': 2,
        'MARÇO': 3,
        'ABRIL': 4,
        'MAIO': 5,
        'JUNHO': 6,
        'JULHO': 7,
        'AGOSTO': 8,
        'SETEMBRO': 9,
        'OUTUBRO': 10,
        'NOVEMBRO': 11,
        'DEZEMBRO': 12
    }

    partes = coluna.split()
    return int(partes[1]), meses_numeros[partes[0]]


def verificar_arquivos_recebidos():
    for mes_ano in os.listdir(base_dir):
        mes_ano_dir = os.path.join(base_dir, mes_ano)

        if os.path.isdir(mes_ano_dir):
            coluna_mes_ano = f'{mes_ano.upper()}'

            if coluna_mes_ano not in df.columns:
                df[coluna_mes_ano] = None  # Adiciona uma nova coluna preenchida com valores nulos

            # Itera sobre os arquivos XML em todas as subpastas
            for arquivo_xml in glob.iglob(os.path.join(mes_ano_dir, '**', '*.xml'), recursive=True):
                # Extrai informações do nome do arquivo
                _, nome_arquivo = os.path.split(arquivo_xml)
                partes_nome = nome_arquivo.split('_')

                if len(partes_nome) >= 3:
                    cnpj, data_recebimento, _ = partes_nome[:3]

                    cnpj_encontrado = df['CNPJ'].astype(str).str.contains(cnpj).any()

                    if coluna_mes_ano in df.columns and cnpj_encontrado:
                        indices_linhas = df[df['CNPJ'].astype(str).str.contains(cnpj)].index.tolist()

                        for indice_linha in indices_linhas:
                            df.at[indice_linha, coluna_mes_ano] = 'OK'

    df['CNPJ'] = df['CNPJ'].astype(str)

    colunas_iniciais = ['CNPJ', 'Nome', 'Casa Externa', 'Status']
    colunas_meses = [coluna for coluna in df.columns if coluna not in colunas_iniciais]

    colunas_ordenadas = sorted(colunas_meses, key=extrair_ano_mes)

    df_organizado = df[colunas_iniciais + colunas_ordenadas]
    df_organizado['CNPJ'] = df_organizado['CNPJ'].astype(str)

    df_organizado.to_excel(caminho_excel, index=False, sheet_name='BASE')


def comparar_bases():
    df1 = pd.read_excel(caminho_excel, sheet_name='BASE').fillna('')
    df2 = excel_copia.fillna('')

    ultima_coluna = df1.columns[-1]
    df_diferencas = df1[df1[ultima_coluna] != df2[ultima_coluna]].copy()

    df_diferencas['Caminho'] = ''

    for indice, linha in df_diferencas.iterrows():
        cnpj = str(linha['CNPJ']).zfill(14)

        for subpasta in os.listdir(os.path.join(base_dir, ultima_coluna)):
            subpasta_path = os.path.join(base_dir, ultima_coluna, subpasta)
            for arquivo in os.listdir(subpasta_path):
                if arquivo.startswith(cnpj):
                    caminho_completo = os.path.join(subpasta_path, arquivo)
                    df_diferencas.at[indice, 'Caminho'] = caminho_completo

    return df_diferencas


def dividir_df():
    ultima_coluna = df.columns[-1]

    df_com_status = df[df['Status'].notnull() & df[ultima_coluna].isnull()]
    df_sem_status = df[df['Status'].isnull() & df[ultima_coluna].isnull()]
    df_com_ok = df[df[ultima_coluna] == 'OK']

    colunas_desejadas = ['CNPJ', 'Nome', 'Casa Externa', 'Status']
    df_com_status = df_com_status[colunas_desejadas].fillna('')
    df_sem_status = df_sem_status[colunas_desejadas].fillna('')
    df_com_ok = df_com_ok[colunas_desejadas].fillna('')

    print("DataFrame com Status:")
    print(df_com_status)
    print("\nDataFrame sem Status:")
    print(df_sem_status)
    print("\nDataFrame com 'OK' na última coluna:")
    print(df_com_ok)


verificar_arquivos_recebidos()
comparar_bases()
dividir_df()
