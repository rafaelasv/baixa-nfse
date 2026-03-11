import openpyxl


def ler_planilha(caminho: str) -> list:
    wb = openpyxl.load_workbook(caminho)
    ws = wb.active
    empresas = []
    for row in ws.iter_rows(min_row=3, values_only=True):  # pula titulo e cabecalho
        nome = str(row[1]).strip() if row[1] else None  # coluna B = Nome
        cnpj = str(row[2]).strip() if row[2] else None  # coluna C = CNPJ
        if nome and cnpj:
            empresas.append({"nome": nome, "cnpj": cnpj})
    return empresas
