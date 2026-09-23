from pathlib import Path
from datetime import date
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import win32com.client as win32

# ============================================================
# CONFIGURAÇÕES DE PASTA
# ============================================================
PASTA_RAIZ = Path(r"X:\Comum-PagamentosDiarios")
MESES = [
    "01 - JANEIRO",
    "02 - FEVEREIRO",
    "03 - MARÇO",
    "04 - ABRIL",
    "05 - MAIO",
    "06 - JUNHO",
    "07 - JULHO",
    "08 - AGOSTO",
    "09 - SETEMBRO",
    "10 - OUTUBRO",
    "11 - NOVEMBRO",
    "12 - DEZEMBRO"
]

# ============================================================
# DATA DA PASTA
# ============================================================
# TEMPORÁRIO PARA TESTES
data_atual = date(2026, 9, 24)
pasta_mes = MESES[data_atual.month - 1]
data_formatada = data_atual.strftime("%d.%m.%Y")

pasta_manha = (
    PASTA_RAIZ
    / str(data_atual.year)
    / pasta_mes
    / data_formatada
    / "MANHÃ"
)
# IDENTIFICAÇÃO DO EXCEL NA PASTA "MANHA"
arquivos_excel = [
    arquivo
    for arquivo in pasta_manha.glob("EXPORT_*.xlsx")
]

print(f"Pasta analisada:")
print(pasta_manha)
print()
print("Arquivos encontrados:")

for arquivo in arquivos_excel:
    print(arquivo.name)

arquivo_1010 = None
arquivo_1013 = None

for arquivo in arquivos_excel:
    planilha = load_workbook(
        arquivo,
        read_only=True,
        data_only=True
    ).active

    empresa = str(planilha["C4"].value).strip()

    print(f"{arquivo.name} → Empresa: {empresa}")

    if empresa == "1010":
        arquivo_1010 = arquivo

    elif empresa == "1013":
        arquivo_1013 = arquivo

print()
print("Arquivo 1010:", arquivo_1010.name if arquivo_1010 else "Não encontrado")
print("Arquivo 1013:", arquivo_1013.name if arquivo_1013 else "Não encontrado")

print()

planilha_1010 = load_workbook(
    arquivo_1010,
    read_only=True,
    data_only=True
).active

planilha_1013 = load_workbook(
    arquivo_1013,
    read_only=True,
    data_only=True
).active

print("Estrutura do arquivo 1010:")
print(f"Linhas: {planilha_1010.max_row}")
print(f"Colunas: {planilha_1010.max_column}")

print()

print("Estrutura do arquivo 1013:")
print(f"Linhas: {planilha_1013.max_row}")
print(f"Colunas: {planilha_1013.max_column}")    
# ============================================================
# COLUNAS DO EXCEL FINAL
# ============================================================

COLUNAS_FINAIS = [
    "Status",
    "Empresa",
    "Parceiro de Negócios",
    "Conta Contrato",
    "Objeto de Seguros",
    "Documento",
    "CPF-CNPJ",
    "Nome",
    "Tipo de Documento",
    "Origem",
    "Data do Documento",
    "Data de Lançamento",
    "Data de Vencimento",
    "Forma de Pagamento",
    "N° Lote de Pagamento",
    "Valor",
    "Banco de Pagamento",
    "Usuário Criação",
    "Objeto de bloqueio",
    "Categoria de bloqueio",
    "Processo de bloqueio",
    "Motivo do bloqueio",
    "Número do Voucher",
    "Banco Empresa",
    "Mensagem"
]

# ============================================================
# CRIAR EXCEL 1010
# ============================================================
novo_arquivo = Workbook()
nova_planilha = novo_arquivo.active
nova_planilha.title = "Dados"

# Cabeçalho
for coluna, titulo in enumerate(COLUNAS_FINAIS, start=1):
    nova_planilha.cell(
        row=1,
        column=coluna,
        value=titulo
    )

# Copiar dados
linha_destino = 2

for linha in range(4, planilha_1010.max_row + 1):
    for coluna_destino, coluna_origem in enumerate(range(2, 27), start=1):
        valor = planilha_1010.cell(
            row=linha,
            column=coluna_origem
        ).value

# TIRA O NEGATIVO DO VALOR
        if coluna_origem == 17 and isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
            valor = abs(float(valor))

# REMOVE ESPAÇOS COLUNA EMPRESA 
        if coluna_origem == 25 and isinstance(valor, str):
            valor = valor.strip()

        nova_planilha.cell(
            row=linha_destino,
            column=coluna_destino,
            value=valor
        )

    linha_destino += 1
# FORMATAR COLUNA VALOR COMO CONTABIL    
for linha in range(2, linha_destino):
    nova_planilha.cell(
        row=linha,
        column=16
    ).number_format = '_-[$R$-pt-BR]* #,##0.00_-;_-[$R$-pt-BR]* -#,##0.00_-;_-[$R$-pt-BR]* "-"??_-;_-@_-'
    
# ATIVAÇÃO DE FILTROS
nova_planilha.auto_filter.ref = f"A1:Y{linha_destino - 1}"

# AJUSTAR COLUNAS
for coluna in range(3, 9):
    nova_planilha.column_dimensions[get_column_letter(coluna)].width = 18
nova_planilha.column_dimensions["P"].width = 18

ARQUIVO_DESTINO = pasta_manha / "TESTE_1010.xlsx"
novo_arquivo.save(ARQUIVO_DESTINO)

print()
print("Arquivo 1010 criado:")
print(ARQUIVO_DESTINO)

# ============================================================
# CRIAR TABELA DINÂMICA 1010
# ============================================================

excel = win32.Dispatch("Excel.Application")
excel.Visible = False

arquivo_excel = excel.Workbooks.Open(str(ARQUIVO_DESTINO))
planilha_excel = arquivo_excel.Worksheets("Dados")

ultima_linha = planilha_excel.Cells(
    planilha_excel.Rows.Count, 1
).End(-4162).Row

# Onde a tabela dinâmica vai começar
linha_dinamica = ultima_linha + 3

# Título acima da tabela dinâmica
titulo = planilha_excel.Cells(
    linha_dinamica,
    1
)

titulo.Value = "RELATÓRIO SAP"
titulo.Font.Bold = True
titulo.Font.Color = 255

# Cria o intervalo de origem da tabela dinâmica
origem = planilha_excel.Range(
    f"A1:Y{ultima_linha}"
)

# Cria o cache da tabela dinâmica
cache = arquivo_excel.PivotCaches().Create(
    SourceType=1,
    SourceData=origem.Address
)

# Define onde a tabela dinâmica será criada
destino = planilha_excel.Cells(
    linha_dinamica + 1,
    1
)

# Cria a tabela dinâmica
tabela_dinamica = cache.CreatePivotTable(
    TableDestination=destino,
    TableName="TabelaDinamica1010"
)

# N° Lote de Pagamento → Linhas
campo_lote = tabela_dinamica.PivotFields(
    "N° Lote de Pagamento"
)
campo_lote.Orientation = 1
campo_lote.Position = 1

# Valor → Valores
campo_valor = tabela_dinamica.AddDataField(
    tabela_dinamica.PivotFields("Valor"),
    "Soma de Valor",
    -4157
)

# Documento → Valores / Contagem
campo_documento = tabela_dinamica.AddDataField(
    tabela_dinamica.PivotFields("Documento"),
    "Contagem de Documento",
    -4112
)

arquivo_excel.Save()
arquivo_excel.Close()
excel.Quit()

print("Tabela dinâmica 1010 criada com sucesso!")
# ============================================================
# CRIAR EXCEL 1013
# ============================================================
novo_arquivo = Workbook()
nova_planilha = novo_arquivo.active
nova_planilha.title = "Dados"

# Cabeçalho
for coluna, titulo in enumerate(COLUNAS_FINAIS, start=1):
    nova_planilha.cell(
        row=1,
        column=coluna,
        value=titulo
    )

# Copiar dados
linha_destino = 2

for linha in range(4, planilha_1013.max_row + 1):
    for coluna_destino, coluna_origem in enumerate(range(2, 27), start=1):
        valor = planilha_1013.cell(
            row=linha,
            column=coluna_origem
        ).value

# TIRA O NEGATIVO DO VALOR
        if coluna_origem == 17 and isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
            valor = abs(float(valor))

# REMOVE ESPAÇOS COLUNA EMPRESA 
        if coluna_origem == 25 and isinstance(valor, str):
            valor = valor.strip()

        nova_planilha.cell(
            row=linha_destino,
            column=coluna_destino,
            value=valor
        )

    linha_destino += 1
    
# FORMATAR COLUNA VALOR COMO CONTABIL    
for linha in range(2, linha_destino):
    nova_planilha.cell(
        row=linha,
        column=16
    ).number_format = '_-[$R$-pt-BR]* #,##0.00_-;_-[$R$-pt-BR]* -#,##0.00_-;_-[$R$-pt-BR]* "-"??_-;_-@_-'
    
# ATIVAÇÃO DE FILTROS
nova_planilha.auto_filter.ref = f"A1:Y{linha_destino - 1}"

# AJUSTAR COLUNAS
for coluna in range(3, 9):
    nova_planilha.column_dimensions[get_column_letter(coluna)].width = 18
nova_planilha.column_dimensions["P"].width = 18

ARQUIVO_DESTINO = pasta_manha / "TESTE_1013.xlsx"
novo_arquivo.save(ARQUIVO_DESTINO)

print()
print("Arquivo 1013 criado:")
print(ARQUIVO_DESTINO)