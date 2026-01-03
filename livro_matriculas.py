import streamlit as st
import pandas as pd
from datetime import date, datetime
import re

# Configuração da página
st.set_page_config(
    page_title="Gerador de Livro de Matrículas",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilo customizado
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True
)

# Utils
def ultimo_dia_maio(ano):
    """Retorna a última quarta-feira do mês de maio do ano informado."""
    dt = date(ano, 5, 31)
    while dt.weekday() != 2: # 2 é quarta-feira
        dt = dt.replace(day=dt.day - 1)
    return dt

def formatar_telefone(telefone):
    """Aplica a máscara (22) XXXXX-XXXX"""
    nums = re.sub(r'\D', '', telefone)
    if not nums: return ""
    if len(nums) == 11:
        return f"({nums[:2]}) {nums[2:7]}-{nums[7:]}"
    elif len(nums) == 10:
        return f"({nums[:2]}) {nums[2:6]}-{nums[6:]}"
    return telefone

def formatar_cep(cep):
    """Aplica a máscara XXXXX-XXX"""
    nums = re.sub(r'\D', '', cep)
    if not nums: return ""
    if len(nums) >= 5:
        return f"{nums[:5]}-{nums[5:8]}"
    return nums

# Header
st.title("Gerador de Livro de Matrículas")
st.markdown("Converta sua planilha gerada no Suap no Livro de Matŕiculas!")
st.write("---")

# Painel de Dados (Sem st.form para permitir reatividade imediata)
st.header("Dados da Unidade Escolar")
st.info("Os dados informados são obrigatórios e serão usados na criação do documento")

# Container para agrupar visualmente
with st.container():
    # Ano Letivo e Data de Referência
    col1, col2 = st.columns(2)
    with col1:
        ano_atual = 2025
        
        # Callback para atualizar a data do censo quando o ano mudar
        def atualizar_data_censo():
            # Pega o valor novo do ano diretamente do estado do widget
            novo_ano = st.session_state.ano_letivo_input
            nova_data = ultimo_dia_maio(int(novo_ano))
            # Atualiza diretamente a chave do widget de data
            st.session_state['data_censo_picker'] = nova_data

        ano_letivo = st.number_input(
            "Ano Letivo",
            min_value=2024, 
            value=max(2025, ano_atual), 
            step=1, 
            format="%d",
            key="ano_letivo_input",
            on_change=atualizar_data_censo
        )
        
        ofertou_eja_1 = st.checkbox("Ofertou EJA no 1º SEM")
        ofertou_eja_2 = st.checkbox("Ofertou EJA no 2º SEM")
    
    with col2:
        # Inicialização garantia da data correta para o ano inicial
        if 'data_censo_picker' not in st.session_state:
            st.session_state['data_censo_picker'] = ultimo_dia_maio(int(ano_letivo))

        data_ref = st.date_input(
            "Data de referência do Censo Escolar",
            format="DD/MM/YYYY",
            key="data_censo_picker" 
        )




    st.write("") # Espaçamento

    # Nome e INEP
    col3, col4 = st.columns([3, 1])
    with col3:
        nome_escola = st.text_input("Nome da Unidade Escolar")
    with col4:
        # Validar 8 dígitos
        codigo_inep = st.text_input("Código do Inep (8 dígitos)", max_chars=8)
        if codigo_inep and (len(codigo_inep) != 8 or not codigo_inep.isdigit()):
            st.warning("O código INEP deve conter exatamente 8 dígitos numéricos.")

    # Endereço (Refatorado)
    st.subheader("Logradouro")
    col_end1, col_end2, col_end3, col_end4 = st.columns([3, 1, 2, 1])
    with col_end1:
        endereco = st.text_input("Endereço")
    with col_end2:
        numero = st.text_input("Nº")
    with col_end3:
        bairro = st.text_input("Bairro")
    with col_end4:
        cep_input = st.text_input("CEP (apenas números)", max_chars=8)
        cep_formatado = ""
        if cep_input:
            if len(cep_input) == 8 and cep_input.isdigit():
                cep_formatado = f"{cep_input[:5]}-{cep_input[5:]}"
                st.caption(f"**CEP Formatado:** {cep_formatado}")
            else:
                st.warning("Digite 8 números")
            
    # Email e Telefone
    col5, col6 = st.columns(2)
    with col5:
        email = st.text_input("E-mail")
        if email and "@" not in email:
            st.error("E-mail inválido")
            
    with col6:
        telefone_input = st.text_input("Telefone (apenas números)", max_chars=11)
        if telefone_input:
            st.caption(f"Formato: {formatar_telefone(telefone_input)}")

    st.write("") # Espaçamento
    st.write("---")

    # Dias Letivos (Alterado)
    col7, col8 = st.columns(2)
    with col7:
        st.markdown("**Turmas Regulares**")
        # Range 200 a 210
        total_dias_letivos = st.number_input("Total de Dias Letivos", min_value=200, max_value=210, value=200, step=1)
    with col8:
        st.markdown("\n.")
       
        data_encerramento = st.date_input("Data de Encerramento", format="DD/MM/YYYY")

    # Condicional EJA 1
    if ofertou_eja_1:
        st.markdown("**EJA 1º Semestre**")
        col_eja1_1, col_eja1_2 = st.columns(2)
        with col_eja1_1:
            dias_eja1 = st.number_input("Total de Dias Letivos - EJA 1º SEM", min_value=100, max_value=110, value=100, step=1)
        with col_eja1_2:
            data_enc_eja1 = st.date_input("Data de Encerramento - EJA 1º SEM", format="DD/MM/YYYY", key="eja1_date")

    # Condicional EJA 2
    if ofertou_eja_2:
        st.markdown("**EJA 2º Semestre**")
        col_eja2_1, col_eja2_2 = st.columns(2)
        with col_eja2_1:
            dias_eja2 = st.number_input("Total de Dias Letivos - EJA 2º SEM", min_value=100, max_value=110, value=100, step=1)
        with col_eja2_2:
            data_enc_eja2 = st.date_input("Data de Encerramento - EJA 2º SEM", format="DD/MM/YYYY", key="eja2_date")

st.write("---")

# Validação Global para Bloquear Botões ou Apenas Avisar
# Como não temos botão "Confirmar", validamos "on the fly"
# Mas o usuário pediu "Abaixo do formulário local para upload"

# Validação Reutilizável
def validar_dados():
    erros = []
    if not nome_escola: erros.append("Nome da escola é obrigatório")
    if not codigo_inep or len(codigo_inep) != 8: erros.append("Confira o código do INEP")
    if not endereco: erros.append("Endereço (logradouro) é obrigatório")
    if not numero: erros.append("Número do endereço é obrigatório")
    if not bairro: erros.append("Bairro é obrigatório")
    if not cep_input or len(cep_input) != 8: erros.append("CEP obrigatório e deve conter 8 dígitos")
    if not email: erros.append("E-mail é obrigatório")
    if not telefone_input: erros.append("Telefone é obrigatório")
    return erros

# Processamento de Dados
@st.cache_data
def carregar_depara():
    try:
        return pd.read_csv("DEPARA.csv")
    except Exception as e:
        st.error(f"Erro ao carregar DEPARA.csv: {e}")
        return None

def tratar_dados(df, ano_letivo_ref, data_censo_ref):
    # 1. Converter datas para DD/MM/yyyy e criar temp para cálculos
    colunas_datas = ["Data de Matrícula", "Data do Último Procedimento"]
    
    # Temp para Data de Matrícula
    if "Data de Matrícula" in df.columns:
        df["_dt_matricula_temp"] = pd.to_datetime(df["Data de Matrícula"], errors='coerce', dayfirst=True)

    for col in colunas_datas:
        if col in df.columns:
            # Garante dayfirst=True para evitar warnings e erros com datas DD/MM
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    
    # 2. Calcular Idade em 31/03 do Ano Letivo
    if "Data de Nascimento" in df.columns:
        df["Data de Nascimento"] = pd.to_datetime(df["Data de Nascimento"], errors='coerce', dayfirst=True)
        
        def calcular_idade(nascimento, ano_ref):
            if pd.isnull(nascimento): return None
            data_ref = date(int(ano_ref), 3, 31)
            idade = data_ref.year - nascimento.year - ((data_ref.month, data_ref.day) < (nascimento.month, nascimento.day))
            return idade
            
        df[f"Idade em 31/03/{ano_letivo_ref}"] = df["Data de Nascimento"].apply(lambda x: calcular_idade(x, ano_letivo_ref))
        
        # Formatar Nascimento para visualização
        df["Data de Nascimento"] = df["Data de Nascimento"].dt.strftime('%d/%m/%Y')

    # 3. Aplicar DE-PARA (Grupo/Ano/Fase e Ordenador)
    df_depara = carregar_depara()
    if df_depara is not None:
        # Chaves para o merge
        chaves = ["Descrição do Curso", "Período no Ano Selecionado"]
        
        # Verificar se as chaves existem no dataframe carregado
        if all(key in df.columns for key in chaves):
            # Garantir tipos compatíveis para merge (convertendo para string se necessário, mas geralmente int/str combinam se limpos)
            # O DEPARA parece ter Período como int. Vamos garantir que no df também seja ou ambos str.
            # Convertendo ambos para string para segurança
            df["Descrição do Curso"] = df["Descrição do Curso"].astype(str)
            df["Período no Ano Selecionado"] = df["Período no Ano Selecionado"].astype(str)
            
            df_depara["Descrição do Curso"] = df_depara["Descrição do Curso"].astype(str)
            df_depara["Período no Ano Selecionado"] = df_depara["Período no Ano Selecionado"].astype(str)

            # Merge Left para manter dados originais
            df = pd.merge(df, df_depara, on=chaves, how="left")
        else:
            st.warning("Colunas 'Descrição do Curso' e 'Período no Ano Selecionado' não encontradas para aplicação do DE-PARA.")

    # 4. Criar Coluna Consolidada de Necessidades Especiais
    # Colunas origem: 'Deficiência', 'Superdotação', 'Transtorno'
    cols_nec = ['Deficiência', 'Superdotação', 'Transtorno']
    
    # 5. Limpeza e Consulta de Naturalidade (Município -> Município(UF))
    if 'Naturalidade' in df.columns:
        # Carregar municípios
        try:
            df_muni = pd.read_csv('municipios.csv')
            # Criar dicionário {Município: UF}
            # Atenção: Municípios homônimos em UFs diferentes (ex: Bom Jesus) vão sobrescrever.
            # O ideal seria ter UF na entrada, mas dada a regra de limpeza anterior, assumimos lookup direto.
            dict_muni = dict(zip(df_muni['Município'], df_muni['UF']))
        except Exception as e:
            st.warning(f"Não foi possível carregar municipios.csv: {e}")
            dict_muni = {}

        def tratar_naturalidade(valor):
            if pd.isnull(valor): return ""
            # Limpeza inicial: Remover tudo após '(' e trim
            nome_limpo = str(valor).split('(', 1)[0].strip()
            
            # Lookup
            if nome_limpo in dict_muni:
                uf = dict_muni[nome_limpo]
                return f"{nome_limpo}({uf})"
            return nome_limpo

        df['Naturalidade'] = df['Naturalidade'].apply(tratar_naturalidade)

    def consolidar_necessidades(row):
        for col in cols_nec:
            if col in row.index:
                valor = row[col]
                # Verifica se não é nulo e não é '-'
                if pd.notna(valor) and str(valor).strip() != '-':
                    return "Sim"
        
        return "-"

    # Aplica linha a linha
    df['Deficiência, TEA, Altas Habilidades ou Superdotação'] = df.apply(consolidar_necessidades, axis=1)

    # Nova Lógica Condicional para 'Situação no Ano Selecionado' (Educação Infantil)
    # Se "Descrição do Curso" == "Educação Infantil" AND "Situação no Ano Selecionado" == "Aprovado" -> "Sem Movimentação"
    # Se "Descrição do Curso" == "Educação Infantil" AND "Situação no Ano Selecionado" == "Reprovado" -> "Ajuste de Idade"
    
    if 'Descrição do Curso' in df.columns and 'Situação no Ano Selecionado' in df.columns:
        # Normalizar strings para comparação segura
        df['Descrição do Curso'] = df['Descrição do Curso'].astype(str).str.strip()
        df['Situação no Ano Selecionado'] = df['Situação no Ano Selecionado'].astype(str).str.strip()
        
        mask_infantil = df['Descrição do Curso'] == 'Educação Infantil'
        mask_aprovado = df['Situação no Ano Selecionado'] == 'Aprovado'
        mask_reprovado = df['Situação no Ano Selecionado'] == 'Reprovado'
        
        df.loc[mask_infantil & mask_aprovado, 'Situação no Ano Selecionado'] = 'Sem Movimentação'
        df.loc[mask_infantil & mask_reprovado, 'Situação no Ano Selecionado'] = 'Ajuste de Idade'

    # Substituição específica solicitada: "Aprovado com Progressão Parcial" -> "Aprovado com Prog. Parcial"
    if 'Situação no Ano Selecionado' in df.columns:
        df['Situação no Ano Selecionado'] = df['Situação no Ano Selecionado'].replace(
            "Aprovado com Progressão Parcial", "Aprovado com Prog. Parcial"
        )
    
    # 6. Lógica Condicional para 'Data da Situação'
    # Mostrar data apenas se Situação for "Transferido" ou "Transf. Externa"
    if 'Situação no Ano Selecionado' in df.columns:
        # Limpar espaços em branco da coluna Situação
        df['Situação no Ano Selecionado'] = df['Situação no Ano Selecionado'].astype(str).str.strip()
        
        if 'Data do Último Procedimento' in df.columns:
            def tratar_data_situacao(row):
                situacao = str(row.get('Situação no Ano Selecionado', '')).strip()
                if situacao in ['Transferido', 'Transf. Externa']:
                    data = row.get('Data do Último Procedimento', '')
                    return data if pd.notna(data) else '-'
                return '-'
            
            df['Data do Último Procedimento'] = df.apply(tratar_data_situacao, axis=1)
        
    # 7. Criar Coluna "Pós Censo"
    # Preenchido com "Sim" se 'Data de Matrícula' (usando temp) for igual ou superior a 'Data de referência do Censo Escolar'
    if "_dt_matricula_temp" in df.columns and data_censo_ref:
        try:
            # Garantir que data_censo_ref seja Timestamp para comparação
            censo_ts = pd.Timestamp(data_censo_ref)
            
            df['Pós Censo'] = df['_dt_matricula_temp'].apply(
                lambda x: "Sim" if pd.notnull(x) and x >= censo_ts else "-"
            )
        except Exception as e:
            st.warning(f"Erro ao calcular Pós Censo: {e}")
            df['Pós Censo'] = "-"
            
        # Remover coluna temporária
        df.drop(columns=["_dt_matricula_temp"], inplace=True)
    else:
         df['Pós Censo'] = "-"

    return df

from pdf_generator import gerar_pdf_matricula, gerar_capa, gerar_termo_abertura, gerar_termo_encerramento

# Validação Reutilizável
def validar_dados(dados):
    erros = []
    if not dados.get('nome'): erros.append("Nome da escola é obrigatório")
    if not dados.get('inep') or len(dados.get('inep', '')) != 8: erros.append("Confira o código do INEP")
    # if not dados.get('logradouro'): erros.append("Endereço (logradouro) é obrigatório") 
    return erros

def processar_arquivo_action(df, key_prefix, dados_escola):
    """Função chamada pelo callback do botão de processar"""
    # Validar
    erros = validar_dados(dados_escola)
    if erros:
        st.session_state[f'erros_{key_prefix}'] = erros
        st.session_state[f'sucesso_{key_prefix}'] = False
    else:
        st.session_state[f'erros_{key_prefix}'] = []
        st.session_state[f'sucesso_{key_prefix}'] = True
        
        # Gerar nome do arquivo conforme regras solicitadas
        if "EJA 2º SEM" in key_prefix:
            filename_pdf = f"livro_matricula{dados_escola['ano_letivo']}2SEM.pdf"
        else:
            filename_pdf = f"livro_matricula{dados_escola['ano_letivo']}.pdf"
        
        # Gerar PDF
        pdf_bytes = gerar_pdf_matricula(df, dados_escola, key_prefix)
        st.session_state[f'pdf_bytes_{key_prefix}'] = pdf_bytes
        st.session_state[f'pdf_name_{key_prefix}'] = filename_pdf


def renderizar_ui_processamento(df, titulo_sucesso, dados_escola):
    key_prefix = titulo_sucesso  # Usando título como chave única
    
    #st.subheader("Pré-visualização (10 primeiras linhas)")
    #st.dataframe(df.head(10))
    #st.write("---")
    
    # Botão de Ação (Callback)
    st.button(
        f"Criar Documentos",
        key=f"btn_criar_{key_prefix}",
        on_click=processar_arquivo_action,
        args=(df, key_prefix, dados_escola)
    )
    
    # Feedback Visual
    if f'erros_{key_prefix}' in st.session_state and st.session_state[f'erros_{key_prefix}']:
        for e in st.session_state[f'erros_{key_prefix}']:
            st.error(e)
            
    if f'sucesso_{key_prefix}' in st.session_state and st.session_state[f'sucesso_{key_prefix}']:
        st.success(f"Arquivos gerados com sucesso!")
        
        # Área de Download Persistente (empilhada verticalmente)
        if f'pdf_bytes_{key_prefix}' in st.session_state:
            # Botão PDF
            st.download_button(
                label=f"Baixar {titulo_sucesso} em PDF",
                data=st.session_state[f'pdf_bytes_{key_prefix}'],
                file_name=st.session_state[f'pdf_name_{key_prefix}'],
                mime="application/pdf",
                key=f"dl_pdf_{key_prefix}"
            )
            if "EJA 2º SEM" not in key_prefix:
                st.write("\n")  # Espaçamento
                # Termo de Abertura
                termo_bytes = gerar_termo_abertura(dados_escola)
                st.download_button(
                    label="Baixar Termo de Abertura",
                    data=termo_bytes,
                    file_name=f"Termo de Abertura {dados_escola.get('ano_letivo', '')}.pdf",
                    mime="application/pdf",
                    key=f"dl_termo_{key_prefix}"
                )
                st.write("\n")
                # Termo de Encerramento
                termo_enc_bytes = gerar_termo_encerramento(dados_escola)
                st.download_button(
                    label="Baixar Termo de Encerramento",
                    data=termo_enc_bytes,
                    file_name=f"Termo de Encerramento {dados_escola.get('ano_letivo', '')}.pdf",
                    mime="application/pdf",
                    key=f"dl_termo_enc_{key_prefix}"
                )
                st.write("\n")
                # Capa
                capa_bytes = gerar_capa(dados_escola)
                st.download_button(
                    label="Baixar Capa do Livro de Matrículas",
                    data=capa_bytes,
                    file_name=f"capa_livro_{dados_escola.get('ano_letivo', 'ano')}.pdf",
                    mime="application/pdf",
                    key=f"dl_capa_{key_prefix}"
                )

col_imp1, col_imp2 = st.columns(2)

# Preparar dados da escola para passar aos callbacks
cep_valor = cep_formatado if 'cep_formatado' in locals() else cep_input

dados_escola = {
    "nome": nome_escola,
    "inep": codigo_inep,
    "ano_letivo": ano_letivo,
    "logradouro": endereco,
    "numero": numero,
    "bairro": bairro,
    "cep": cep_valor,
    "telefone": telefone_input,
    "email": email,
    # Novos campos para o cabeçalho dinâmico
    "data_censo": data_ref,
    "data_encerramento": data_encerramento,
    "total_dias_letivos": total_dias_letivos,
    "data_enc_eja1": data_enc_eja1 if 'data_enc_eja1' in locals() else None,
    "dias_eja1": dias_eja1 if 'dias_eja1' in locals() else None,
    "data_enc_eja2": data_enc_eja2 if 'data_enc_eja2' in locals() else None,
    "dias_eja2": dias_eja2 if 'dias_eja2' in locals() else None
}

# COLUNA 1: Importar Arquivo Regular / EJA 1
with col_imp1:
    st.markdown("### Importar Arquivo do SUAP")
    st.info("Use esta opção para turmas Regulares e EJA 1º Semestre.")
    
    uploaded_file1 = st.file_uploader("Upload Regular", type=["xlsx", "xls", "csv"], key="up1")
    
    if uploaded_file1 is not None:
        try:
            if uploaded_file1.name.endswith('.csv'):
                df1 = pd.read_csv(uploaded_file1)
            else:
                df1 = pd.read_excel(uploaded_file1)
            
            df1 = tratar_dados(df1, ano_letivo, dados_escola['data_censo'])
            renderizar_ui_processamento(df1, "Livro de Matrículas", dados_escola)
            
        except Exception as e:
            st.error(f"Erro ao processar arquivo 1: {e}")

# COLUNA 2: Importar Arquivo EJA 2
with col_imp2:
    st.markdown("### Importar Arquivo EJA 2º Sem (EM DESENVOLVIMENTO)")
    st.info("Opção exclusiva para turmas de EJA do 2º Semestre.")
    
    uploaded_file2 = st.file_uploader("Upload EJA 2", type=["xlsx", "xls", "csv"], key="up2")
    
    if uploaded_file2 is not None:
        try:
            if uploaded_file2.name.endswith('.csv'):
                df2 = pd.read_csv(uploaded_file2)
            else:
                df2 = pd.read_excel(uploaded_file2)
            
            df2 = tratar_dados(df2, ano_letivo, dados_escola['data_censo'])
            renderizar_ui_processamento(df2, "Livro EJA 2º SEM", dados_escola)
            
        except Exception as e:
            st.error(f"Erro ao processar arquivo 2: {e}")
