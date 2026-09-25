from pathlib import Path
from docx import Document
from datetime import date
import os

pasta_raiz = Path(os.environ["AUTOMACAO_PASTA_RAIZ"])
data_atual = date.fromisoformat(os.environ["AUTOMACAO_DATA_PASTA"])
pasta_data = (
	pasta_raiz
	/ str(data_atual.year)
	/ [
		"01 - JANEIRO", "02 - FEVEREIRO", "03 - MARÇO",
		"04 - ABRIL", "05 - MAIO", "06 - JUNHO",
		"07 - JULHO", "08 - AGOSTO", "09 - SETEMBRO",
		"10 - OUTUBRO", "11 - NOVEMBRO", "12 - DEZEMBRO"
	][data_atual.month - 1]
	/ data_atual.strftime("%d.%m.%Y")
)
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

renda_is = pasta_manha / "RENDA IS"
renda_is.mkdir(exist_ok=True)
renda_rg = pasta_manha / "RENDA RG"
renda_rg.mkdir(exist_ok=True)

portabilidade_is = pasta_manha / "PORTABILIDADE IS"
portabilidade_is.mkdir(exist_ok=True)
portabilidade_rg = pasta_manha / "PORTABILIDADE RG"
portabilidade_rg.mkdir(exist_ok=True)

lancamentos = pasta_manha / "LANÇAMENTOS MANUAIS"
lancamentos.mkdir(exist_ok=True)