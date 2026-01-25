# Gerador de Livro de Matrículas

Este projeto foi gerado com auxílio de IA na IDE Antigravity e é uma ferramenta desenvolvida em Python com [Streamlit](https://streamlit.io/) para automatizar a criação do Livro de Matrículas escolar. A partir de dados exportados do SUAP (Sistema Unificado de Administração Pública), o sistema gera arquivos PDF formatados prontos para impressão e encadernação.

## Funcionalidades

- **Importação de Dados**: Aceita arquivos `.csv` e `.xlsx` (Excel) exportados do SUAP.
- **Tratamento Automático**:
  - Formatação e máscara de CPF, Telefone e CEP.
  - Cálculo automático da idade do aluno na data de corte (31/03).
  - Deduplicação de registros (mantendo a matrícula mais recente se houver duplicidade de CPF).
  - Identificação de alunos com necessidades especiais (PNE).
- **Geração de Documentos**:
  - **Livro de Matrículas**: Relatório em formato A4 Paisagem com listagem dos alunos.
  - **Capa**: Capa padronizada com brasão e dados da unidade escolar.
  - **Termos**: Gera automaticamente Termo de Abertura e Termo de Encerramento.
- **Suporte a Diferentes Modalidades**:
  - Turmas Regulares
  - EJA (Educação de Jovens e Adultos) - 1º e 2º Semestres.

## Pré-requisitos

- Python 3.12 ou superior.
- [uv](https://github.com/astral-sh/uv) (Recomendado para gerenciamento de dependências) ou `pip`.

## Instalação

### Opção 1: Usando `uv` (Recomendado)

Este projeto já está configurado com `uv`. Para sincronizar o ambiente e instalar as dependências exatas definidas no `uv.lock`, execute:

```bash
uv sync
```

### Opção 2: Usando `pip`

Se preferir utilizar o `pip` tradicional:

1. (Opcional) Crie e ative um ambiente virtual:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Linux/Mac
   # ou .venv\Scripts\activate no Windows
   ```

2. Instale as dependências listadas no `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```

## Como Usar

1. **Inicie a aplicação**:

   Se estiver usando `uv`:
   ```bash
   uv run streamlit run livro_matriculas.py
   ```

   Se estiver usando `pip`:
   ```bash
   streamlit run livro_matriculas.py
   ```

2. **No Navegador**:
   - A interface abrirá automaticamente no seu navegador padrão (geralmente em `http://localhost:8501`).
   - Preencha os **Dados da Unidade Escolar** (Nome, INEP, Endereço, etc).
   - Defina o **Ano Letivo** e **Datas** (Censo, Encerramento).
   - Na seção de upload, escolha o arquivo do SUAP correspondente ("Upload Regular" ou "Upload EJA 2").
   - Clique em **"Criar Documentos"**.
   - Faça o download dos PDFs gerados (Livro, Capa e Termos).

## Estrutura do Projeto

- `livro_matriculas.py`: Código principal da interface Streamlit e lógica de tratamento de dados.
- `pdf_generator.py`: Módulo responsável pela criação dos PDFs usando a biblioteca `fpdf`.
- `DEPARA.csv`: Arquivo csv usado para mapeamento de cursos/matrizes (opcional, se necessário).
- `municipios.csv`: Base de dados para identificação da UF baseada no nome do município.
- `brasao.png`: Imagem do brasão utilizada na capa e cabeçalhos.

## Autoria

Desenvolvido para secretarias de creches e escolas da rede municipal de Campos dos Goytacazes (RJ).
