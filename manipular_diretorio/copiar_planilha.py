from pathlib import Path
from datetime import date, timedelta
import os
import shutil
from abrir_pasta import pasta_manha

# CCONFIGURAÇÕES E IMPORTAÇÕES

NOME_PLANILHA = "Planilha de aprovação_Manhã.xlsx"
PASTA_RAIZ = Path(os.environ["AUTOMACAO_PASTA_ORIGEM"])

# ENCONTRAR A ÚLTIMA PLANILHA UTILIZADA

def encontrar_ultima_planilha():

# COMEÇA PELO DIA ANTERIOR
    data = date.today() - timedelta(days=1)

    meses = [
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

    while True:

        pasta_mes = meses[data.month - 1]
        pasta = (
            PASTA_RAIZ
            / str(data.year)
            / pasta_mes
            / data.strftime("%d.%m.%Y")
            / "MANHÃ"
        )

        arquivo = pasta / NOME_PLANILHA
        print(f"Procurando: {arquivo}")

        if arquivo.exists():

            print()
            print("Planilha encontrada:")
            print(arquivo)
            print()

            return arquivo

    # VOLTA UM DIA
        data -= timedelta(days=1)


# COPIAR PLANILHA
def copiar_planilha():

    origem = encontrar_ultima_planilha()

    destino = pasta_manha / NOME_PLANILHA

    print("Copiando planilha...")
    print()
    print(f"Origem:")
    print(origem)
    print()
    print(f"Destino:")
    print(destino)
    print()

# NÃO SOBREESCREVER CASO EXISTA
    if destino.exists():

        print("A planilha de hoje já existe.")
        print("Nada foi sobrescrito.")

        return

    shutil.copy2(origem, destino)

    print("Planilha copiada com sucesso!")
    print()


# EXECUÇÃO DA FUNÇÃO
copiar_planilha()
