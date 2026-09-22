from pathlib import Path
from docx import Document
from datetime import date

data_atual = date.today().strftime("%d.%m.%Y")
pasta_data = Path(r"X:\Comum-PagamentosDiarios\2026\09 - SETEMBRO") / "26.09.2026"
pasta_data.mkdir(exist_ok=True)

pasta_manha = pasta_data / "MANHÃ"
pasta_manha.mkdir(exist_ok=True)

devolucao_is = pasta_manha / "DEVOLUÇÃO IS"
devolucao_is.mkdir(exist_ok=True)
docword = Document()
docword.save(devolucao_is / "Confirmação de envio de pagamentos - Devolução IS.docx")

devolucao_rg = pasta_manha / "DEVOLUÇÃO RG"
devolucao_rg.mkdir(exist_ok=True)
docword = Document()
docword.save(devolucao_rg / "Confirmação de envio de pagamentos - Devolução RG.docx")

resgate_is = pasta_manha / "RESGATE IS"
resgate_is.mkdir(exist_ok=True)
resgate_rg = pasta_manha / "RESGATE RG"
resgate_rg.mkdir(exist_ok=True)

portabilidade_is = pasta_manha / "PORTABILIDADE IS"
portabilidade_is.mkdir(exist_ok=True)
portabilidade_rg = pasta_manha / "PORTABILIDADE RG"
portabilidade_rg.mkdir(exist_ok=True)

lancamentos = pasta_manha / "LANÇAMENTOS MANUAIS"
lancamentos.mkdir(exist_ok=True)