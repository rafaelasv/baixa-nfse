import os

URL_LOGIN     = "https://www.nfse.gov.br/EmissorNacional/"
URL_RECEBIDAS = "https://www.nfse.gov.br/EmissorNacional/Notas/Recebidas"

TIPO_DOWNLOAD = "XML"  # "XML" ou "PDF"

PASTA_SAIDA_PADRAO = os.path.join(os.path.expanduser("~"), "Downloads", "Downloads_XML")

TIMEOUT_LOGIN = 120  # segundos aguardando o usuario selecionar o certificado
