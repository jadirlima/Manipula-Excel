from datetime import date

from openpyxl.workbook import Workbook

# acao = input("Qual a código da Ação que vc quer processar? ")
acao = "BIDI4"
with open(f'./dados/{acao}.txt', 'r') as arquivo_cotacao:
    linhas = arquivo_cotacao.readlines()
    linhas =[linha.replace("\n","").split(";") for linha in linhas]

workbook = Workbook()
planilha_ativa = workbook.active
planilha_ativa.title = "Dados"

planilha_ativa.append(["DATA","COTAÇÃO","BANDA INFERIOR","BANDA SUPERIOR"])

indice = 2

for linha in linhas:
    # Data
    ano_mes_dia = linha[0].split(" ")[0]
    data = date(
        year=int(ano_mes_dia.split("-")[0]),
        month=int(ano_mes_dia.split("-")[1]),
        day=int(ano_mes_dia.split("-")[2])
    )
    # Cotação
    cotacao = float(linha[1])

    # Atualiza as celulas da Planilha Ativa do Excel
    planilha_ativa[f'A{indice}'] = data
    planilha_ativa[f'B{indice}'] = cotacao
    planilha_ativa[f'C{indice}'] = f'=AVERAGE(B{indice}:B{indice+19}) - 2*STDEV(B{indice}:B{indice+19})'
    planilha_ativa[f'D{indice}'] = f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'

    indice += 1


workbook.save("./saida/Planilha.xlsx")