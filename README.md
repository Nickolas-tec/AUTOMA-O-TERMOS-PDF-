
# Automação de Processos com Python: Gerador de Documentos

![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)

## 🎯 Visão Geral

Este projeto demonstra a construção de um pipeline de ETL (Extração, Transformação e Carga) em Python para automatizar a criação de documentos personalizados (formatos `.docx` e `.pdf`). A solução lê dados de uma planilha Excel, realiza tratamentos e filtros, e utiliza um documento Word como template para gerar arquivos individuais para cada registro que atende aos critérios definidos.

Este é um exemplo prático de como a automação com Python pode otimizar tarefas repetitivas, reduzir erros manuais e aumentar a eficiência em processos de RH, administrativos e financeiros.

## ✨ Funcionalidades Principais

- **Extração e Limpeza de Dados**: Leitura de planilhas `.xlsx` utilizando a biblioteca **Pandas**.
- **Tratamento e Validação**: Conversão e validação de tipos de dados, como datas, e normalização de campos de texto.
- **Filtragem Inteligente**: Seleção de dados com base em critérios de negócio específicos (neste caso, o perfil de cargo).
- **Geração Dinâmica de Documentos**: Preenchimento automático de um template `.docx` com dados de colaboradores, funcionando como uma mala direta programática.
- **Conversão para PDF**: Geração de uma versão em `.pdf` de cada documento Word criado, garantindo um formato final para distribuição.
- **Organização de Arquivos**: Criação automática de uma estrutura de pastas para armazenar os documentos de forma organizada por perfil e colaborador.
- **Robustez**: O script inclui tratamento de erros e lida com dados ausentes (`NaN`) de forma elegante, garantindo que o processo não seja interrompido por falhas em registros individuais.

## ⚙️ Como Funciona: O Fluxo ETL

O processo segue um fluxo de ETL bem definido:

1.  **Extração (Extract)**: O script inicia lendo os dados brutos da planilha `BASE 2025.xlsx`.
2.  **Transformação (Transform)**:
    - Colunas desnecessárias são descartadas.
    - Os dados são limpos (espaços em branco removidos, texto em minúsculas para padronização).
    - As colunas de data são convertidas para o formato `datetime`.
    - Os registros são filtrados com base no valor da coluna `Perfil Contrato`.
    - O resultado do tratamento é salvo em uma nova planilha, `CARGOS_FILTRADOS.xlsx`, para auditoria e clareza do processo.
3.  **Carga (Load)**:
    - O script itera sobre cada linha da planilha filtrada.
    - Para cada linha, ele cria uma pasta para o colaborador.
    - O template `FICUS.docx` é carregado em memória.
    - As "tags" ou "placeholders" (ex: `COLABORADOR`, `000.000.000-00`) são substituídos pelos dados do colaborador.
    - Um novo arquivo `.docx` é salvo na pasta do colaborador.
    - O arquivo `.docx` é convertido para `.pdf` na mesma pasta.

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **Pandas**: Para manipulação e análise de dados.
- **python-docx**: Para criar e manipular arquivos do Microsoft Word (`.docx`).
- **docx2pdf**: Para converter os arquivos `.docx` em `.pdf`.
- **XlsxWriter**: Como engine do Pandas para garantir a formatação correta dos dados no Excel de saída.

## 🚀 Como Executar o Projeto

### Pré-requisitos

- **Python 3.7** ou superior.
- **Microsoft Word** instalado (necessário para a biblioteca `docx2pdf`).
- Arquivos de entrada:
    - `BASE 2025.xlsx`: Planilha com os dados brutos dos colaboradores.
    - `FICUS.docx`: Documento Word servindo como template.

### Passos para Execução

1.  **Clone o repositório:**
    ```bash
    git clone <URL-DO-SEU-REPOSITORIO>
    cd <NOME-DO-SEU-PROJETO>
    ```

2.  **Crie um ambiente virtual (recomendado):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # No Windows: venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Execute o script:**
    ```bash
    python "script 02.py"
    ```

5.  **Verifique a saída:** Após a execução, uma nova pasta será criada no diretório do projeto com o nome do perfil filtrado. Dentro dela, haverá subpastas para cada colaborador contendo os arquivos `.docx` e `.pdf` gerados.

