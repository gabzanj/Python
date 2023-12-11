import os
import random
from datetime import datetime, timedelta
import calendar
import openpyxl
from faker import Faker


def criar_cnpj_aleatorio():
    return ''.join([str(random.randint(0, 9)) for _ in range(14)])


def criar_data_aleatoria(data_inicial, data_final):
    delta = data_final - data_inicial
    dias_aleatorios = random.randint(0, delta.days)
    data_aleatoria = data_inicial + timedelta(days=dias_aleatorios)
    return data_aleatoria


def criar_arquivo_xml_e_adicionar_cnpj(caminho_pasta, cnpj, data_envio, data_recebimento, nome, status, planilha):
    nome_arquivo = f"{cnpj}_{data_envio}_{data_recebimento}_arquivo.xml"
    caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)

    conteudo_xml = f"""<root>
  <cnpj>{cnpj}</cnpj>
  <data_envio>{data_envio}</data_envio>
  <data_recebimento>{data_recebimento}</data_recebimento>
  <casa_externa>{nome}</casa_externa>
  <nome>{nome}</nome>
  <status>{status}</status>
</root>"""

    with open(caminho_arquivo, 'w') as arquivo:
        arquivo.write(conteudo_xml)

    planilha.append([cnpj, nome, nome, status])


def criar_estrutura_diretorios_e_planilha_excel(base_dir):
    meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
             'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

    fake = Faker('pt_BR')

    ano_atual = datetime.now().year

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "BASE"

    sheet.append(["CNPJ", "Nome", "Casa Externa", "Status"])

    for mes in meses:
        mes_dir = os.path.join(base_dir, f"{mes} {ano_atual}")
        os.makedirs(mes_dir, exist_ok=True)

        ultimo_dia_mes = calendar.monthrange(ano_atual, meses.index(mes) + 1)[1]

        for dia in range(1, ultimo_dia_mes + 1):
            dia_str = str(dia).zfill(2)
            enviado_dir = os.path.join(mes_dir, f"Enviado em {dia_str}{mes}{ano_atual}")
            os.makedirs(enviado_dir, exist_ok=True)

            cnpj = criar_cnpj_aleatorio()
            nome = fake.name()
            data_envio = f"{ano_atual}{meses.index(mes)+1:02d}{dia_str}"
            data_recebimento = f"{ano_atual}{meses.index(mes)+1:02d}{dia_str}"

            status = random.choice([None, '60 dias', '90 dias'])

            criar_arquivo_xml_e_adicionar_cnpj(enviado_dir, cnpj, data_envio, data_recebimento, nome, status, sheet)

    workbook.save(os.path.join(base_dir, "cnpjs_com_nomes_status.xlsx"))


if __name__ == "__main__":
    diretorio_base = "Diretorio Base"
    criar_estrutura_diretorios_e_planilha_excel(diretorio_base)
    print("Estrutura de diretórios, arquivos XML e planilha Excel criadas com sucesso.")
