# Automação de Download - NFS-e Nacional

Script Python com interface gráfica para baixar automaticamente notas fiscais recebidas do [Portal NFS-e Nacional](https://www.nfse.gov.br/EmissorNacional/), iterando sobre múltiplas empresas a partir de uma planilha Excel.

---

## Funcionalidades

- Lê uma lista de empresas de uma planilha `.xlsx`
- Abre o navegador automaticamente para cada empresa
- Detecta o login via certificado digital sem intervenção manual
- Navega até **Notas Recebidas** e aplica o filtro de período
- Baixa os arquivos XML (ou DANFS-e em PDF) de todas as páginas
- Salva cada empresa em sua própria pasta `NomeEmpresa_CNPJ`

---

## Pré-requisitos

- Python 3.10 ou superior
- Google Chrome instalado
- Certificado digital A1 ou A3 configurado no navegador

Instale as dependências:

```bash
pip install selenium openpyxl customtkinter pillow
```

> O ChromeDriver é gerenciado automaticamente pelo `selenium-manager` — não é necessário instalar manualmente.

---

## Estrutura da Planilha

O arquivo `.xlsx` deve seguir este formato:

| (linha 1) | Empresas | | |
|-----------|----------|------|------|
| **Codigo** | **Nome** | **CNPJ** | |
| 112 | EMPRESA EXEMPLO LTDA | 00.000.000/0001-00 | |

- Linha 1: título (ignorado)
- Linha 2: cabeçalho (ignorado)
- Linha 3 em diante: dados das empresas

> Um arquivo `template_planilha.xlsx` está disponível na raiz do projeto para uso como base.

---

## Como usar

1. Execute o script:
   ```bash
   python main.py
   ```

2. Na interface, preencha:
   - **Planilha (.xlsx):** selecione o arquivo com a lista de empresas
   - **Salvar em:** pasta onde os XMLs serão salvos
   - **Data Início / Data Fim:** período desejado (máximo 30 dias por consulta)
   - **Tipo:** `xml` para NFS-e em XML ou `pdf` para DANFS-e

3. Clique em **INICIAR PROCESSO**

4. Para cada empresa, o Chrome abrirá automaticamente — selecione o certificado digital quando solicitado. O script detecta o login e continua sozinho.

5. Os arquivos serão salvos em subpastas:
   ```
   Pasta escolhida/
   ├── EMPRESA A_00.000.000_0001-00/
   │   ├── nfse_xyz.xml
   │   └── ...
   ├── EMPRESA B_11.111.111_0001-11/
   │   └── ...
   ```

---

## Observações

- O portal limita o filtro a **no máximo 30 dias** por consulta. Períodos maiores precisam ser divididos em múltiplas execuções.
- O script **não armazena** nem transmite credenciais — o login é feito diretamente pelo usuário no navegador.
- Em caso de falha no filtro de datas, o script pula a empresa e registra o erro no log.

---

## Tecnologias

- [Python](https://www.python.org/)
- [Selenium](https://www.selenium.dev/) — automação do navegador
- [openpyxl](https://openpyxl.readthedocs.io/) — leitura da planilha Excel
- [customtkinter](https://github.com/TomSchimansky/CustomTkinter) — interface gráfica moderna
- [Pillow](https://python-pillow.org/) — carregamento da logo na interface
- [tkinter](https://docs.python.org/3/library/tkinter.html) — diálogos de arquivo