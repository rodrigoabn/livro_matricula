from fpdf import FPDF
from datetime import datetime
import pandas as pd
import math

def fix_text(text):
    if text is None: return ""
    s = str(text)
    s = s.replace('–', '-').replace('—', '-')
    s = s.replace('“', '"').replace('”', '"')
    s = s.replace("‘", "'").replace("’", "'")
    return s.encode('latin-1', 'replace').decode('latin-1')

class PDF(FPDF):
    def __init__(self, titulo_doc, dados_escola):
        super().__init__(orientation='L', unit='mm', format='A4')
        self.titulo_doc = titulo_doc
        self.dados_escola = dados_escola
        # Margens: Left, Top, Right
        self.set_margins(10, 30, 10)  # margem superior 3cm
        self.set_auto_page_break(auto=True, margin=10)
        self.current_header_info = {} # Dicionário para informações dinâmicas do cabeçalho por grupo

    def header(self):
        # Configurar larguras e posições
        margin_left = 5
        page_width = 297  # A4 Landscape
        usable_width = page_width - 2 * margin_left
        
        # Dados Escola
        nome = fix_text(self.dados_escola.get('nome', ''))
        inep = fix_text(self.dados_escola.get('inep', ''))
        ano = fix_text(self.dados_escola.get('ano_letivo', ''))
        end = fix_text(self.dados_escola.get('logradouro', ''))
        num = fix_text(self.dados_escola.get('numero', ''))
        bairro = fix_text(self.dados_escola.get('bairro', ''))
        cep = fix_text(self.dados_escola.get('cep', ''))
        tel = fix_text(self.dados_escola.get('telefone', ''))
        email = fix_text(self.dados_escola.get('email', ''))

        self.set_y(20)  # Margem superior 20mm
        self.set_x(margin_left)
        
        # Definir posições iniciais para uso posterior
        x_start = self.get_x()
        y_start = self.get_y()
        
# --- Linha 1: Logo (Col 1) | Info Institucional (Col 2) | Dados da Escola (Col 3) ---
        h_row1 = 20
        w_col1 = 25
        w_col2 = (usable_width - w_col1) / 2
        w_col3 = w_col2

        # Coluna 1: Logo
        # self.rect(x_start, y_start, w_col1, h_row1)
        logo_height = 12.8 # Reduzido em 20%
        logo_width = 12.8  # Reduzido em 20%
        if logo_height > h_row1: logo_height = h_row1 - 2
        
        # Centralizado horizontalmente, alinhado ao topo
        # logo_padding = (h_row1 - logo_height) / 2 # REMOVIDO para alinhar ao topo
        
        try:
             # Alinhando ao topo com padding top = 2mm (igual Col 2)
             self.image('brasao.png', x_start + (w_col1 - logo_width)/2, y_start + 2, logo_width, logo_height)
        except:
             pass

        # Coluna 2: Informações Institucionais
        self.set_xy(x_start + w_col1, y_start)
        # self.rect(x_start + w_col1, y_start, w_col2, h_row1)
        self.set_font('Arial', '', 6)
        texto_inst = [
            "Estado do Rio de Janeiro",
            "Prefeitura Municipal de Campos dos Goytacazes",
            "Secretaria Municipal de Educação, Ciência e Tecnologia",
            "Diretoria de Supervisão Escolar"
        ]
        cur_x2 = x_start + w_col1 + 2
        cur_y2 = y_start + 2  # Padding aqui é 2
        line_h2 = 4
        for linha in texto_inst:
            self.set_xy(cur_x2, cur_y2)
            self.cell(w_col2 - 4, line_h2, linha, 0, 0, 'L')
            cur_y2 += line_h2

        # Coluna 3: Dados da Escola (alinhado à direita)
        self.set_xy(x_start + w_col1 + w_col2 , y_start)
        # self.rect(x_start + w_col1 + w_col2, y_start, w_col3, h_row1)
        self.set_font('Arial', 'B', 10)
        linha1 = f"{nome}"
        self.set_font('Arial', '', 6)
        # Formatar telefone
        tel_nums = "".join(filter(str.isdigit, str(tel)))
        if len(tel_nums) == 11:
            tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:7]}-{tel_nums[7:]}"
        elif len(tel_nums) == 10:
            tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:6]}-{tel_nums[6:]}"
        else:
            tel_fmt = tel

        linha2 = f"Endereço: {end} , {num}. {bairro}. CEP: {cep}"
        linha3 = f"Contatos: {tel_fmt} | {email}"
        linha4 = f"Código do INEP: {inep}"
        cur_x3 = x_start + w_col1 + w_col2
        cur_y3 = y_start + 2
        line_h3 = 4
        
        # Linha 1: Negrito 10
        self.set_font('Arial', 'B', 10)
        self.set_xy(cur_x3, cur_y3)
        self.cell(w_col3, line_h3, linha1, 0, 0, 'R')
        cur_y3 += line_h3
        
        # Demais linhas: Regular 6
        self.set_font('Arial', '', 6)
        for txt in [linha2, linha3, linha4]:
            self.set_xy(cur_x3, cur_y3)
            self.cell(w_col3, line_h3, txt, 0, 0, 'R')
            cur_y3 += line_h3

        # --- Linha 2: Título (centralizado) ---
        h_row2 = 5
        y_row2 = y_start + h_row1 - 1
        self.set_xy(x_start, y_row2)
        # self.rect(x_start, y_row2, usable_width, h_row2)
        
        self.set_font('Arial', 'B', 12)
        # Lógica de Sufixo do Título baseada no botão e curso
        suffix = ""
        is_btn_2 = "EJA 2º SEM" in self.titulo_doc
        
        # A informação do curso está em current_header_info (populado antes do add_page)
        desc_curso = self.current_header_info.get('descricao_curso', '')
        # Verificar se o curso corresponde a alguma fase EJA (Iniciais ou Finais)
        eja_phases = ['Educação de Jovens e Adultos Fases Iniciais', 'Educação de Jovens e Adultos Fases Finais']
        is_eja_phase = any(phase in desc_curso for phase in eja_phases)
        
        if is_btn_2:
            suffix = "(EJA - 2º SEM)"
        elif is_eja_phase:
            suffix = "(EJA - 1º SEM)"
        else:
            suffix = ""
    
        titulo_formatado = f"Livro de Matrículas {ano}{suffix}"
      
        x_start1 = (x_start + 2)
        self.set_xy(x_start1, y_row2) # Retirado o + 1.5 para alinhar ao topo
        self.cell(usable_width, 1, titulo_formatado, 0, 0, 'C')

        # --- Linha 3: Informações Dinâmicas (alinhado à esquerda) ---
        h_row3 = 5
        y_row3 = y_row2 + h_row2
        
        
        self.set_xy(x_start1, y_row3)
        self.set_font('Arial', '', 7)  # Negrito para os rótulos
        
        # Recuperar informações do grupo atual (ou usar vazio se não definido)
        info = self.current_header_info
        
        # Montar texto formatado
        # Usando multiplas células lado a lado para espaçamento controlado ou string única?
        # "colocando as informações lado a lado , com espaçamento entre elas"
        
        # Labels e Valores
        turma_txt = f"Turma: {fix_text(info.get('turma', ''))}"
        matriz_txt = f"Matriz: {fix_text(info.get('matriz', ''))}"
        
        # Formatar datas se existirem e forem objetos date/datetime
        dt_censo = info.get('data_censo', '')
        if hasattr(dt_censo, 'strftime'): dt_censo = dt_censo.strftime('%d/%m/%Y')
        censo_txt = f"Data Base do Censo: {dt_censo}"
        
        dt_enc = info.get('data_encerramento', '')
        if hasattr(dt_enc, 'strftime'): dt_enc = dt_enc.strftime('%d/%m/%Y')
        enc_txt = f"Data de Encerramento: {dt_enc}"
        
        dias_txt = f"Dias Letivos: {fix_text(info.get('dias_letivos', ''))}"
        
        # Espaçamento manual com string única (mais seguro para alignment L)
        spacer = "   |   "
        
        # Coluna 1 (Esquerda)
        text_left = f"{turma_txt}{spacer}{matriz_txt}"
        
        # Coluna 2 (Direita)
        text_right = f"{censo_txt}{spacer}{enc_txt}{spacer}{dias_txt}"
        
        # Se for muito longo, reduzir fonte? (Checando ambas as strings somadas + margem)
        if self.get_string_width(text_left + "   " + text_right) > usable_width:
             self.set_font('Arial', '', 7)
        
        # Ajuste vertical
        # Desenhar Coluna Esquerda
        self.set_xy(x_start + 5, y_row3 + 1)
        self.cell(usable_width, 3, text_left, 0, 0, 'L')
        
        # Desenhar Coluna Direita (Sobreposta na mesma linha, alinhada à direita)
        self.set_xy(x_start, y_row3 + 1)
        self.cell(usable_width - 2, 3, text_right, 0, 0, 'R')
        
        # Espaço final entre cabeçalho e tabela
        # Ajustando posição Y final para garantir espaçamento antes da tabela
        self.set_y(y_row3 + h_row3 + 4)

    def footer(self):
        # Posição do rodapé (Mantendo original)
        self.set_y(-8)
        
        # Linha de assinaturas
        self.set_font('Arial', '', 6)
        
        # Calcular largura disponível e posições
        page_width = self.w
        margin = self.l_margin
        usable_width = page_width - 2 * margin
        
        # Duas colunas para assinaturas
        col_width = usable_width / 2
        
        # Primeira coluna - Direção
        x_start = margin
        self.set_xy(x_start, self.get_y())
        
        # Linha para assinatura (60mm de largura)
        line_width = 60
        x_line1 = x_start + (col_width - line_width) / 2
        y_line = self.get_y()
        self.line(x_line1, y_line, x_line1 + line_width, y_line)
        
        # Texto "Direção" abaixo da linha
        self.set_xy(x_start, y_line + 1)
        self.cell(col_width, 3, 'Direção', 0, 0, 'C')
        
        # Segunda coluna - Supervisor Pedagogo
        x_start2 = margin + col_width
        x_line2 = x_start2 + (col_width - line_width) / 2
        self.line(x_line2, y_line, x_line2 + line_width, y_line)
        
        # Texto "Supervisor Pedagogo"
        self.set_xy(x_start2, y_line + 1)
        self.cell(col_width, 3, 'Supervisor Pedagogo', 0, 0, 'C')
        

def gerar_pdf_matricula(df, dados_escola, titulo_documento):
    pdf = PDF(titulo_documento, dados_escola)
    # Configuração de Fonte reduzida para caber muitas colunas
    pdf.set_font('Arial', 'B', 6)

    # 1. Preparar Dados para o Relatório
    df_relatorio = df.copy()
    
    # Criar coluna de Ordem Sequencial com nome "#"
    # Garante que não duplica se já existir no arquivo original
    if '#' in df_relatorio.columns:
        df_relatorio.drop(columns=['#'], inplace=True)
        
    # df_relatorio.insert(0, '#', range(1, len(df_relatorio) + 1)) # Moved to inside group loop
    
    # Mapeamento de Colunas
    ano_letivo = dados_escola.get('ano_letivo', '')
    col_age_dynamic = f"Idade em 31/03/{ano_letivo}"
    
    # Ajustando para garantir que a chave '#' seja usada consistentemente
    mapa_colunas = {
        '#': '#', 
        'Grupo/Ano/Fase': 'Grupo/Ano/Fase',
        'Matrícula': 'Matrícula',
        'CPF': 'CPF',
        'Nome': 'Nome',
        'Data de Nascimento': 'Data de Nascimento',
        col_age_dynamic: 'Idade (31/03)',
        'Sexo': 'Sexo',
        'Nome da Mãe': 'Filiação 1',
        'Nome do Pai': 'Filiação 2',
        'Naturalidade': 'Naturalidade',
        'Nacionalidade': 'Nacionalidade',
        'Data de Matrícula': 'Data Matrícula Suap',
        'Deficiência, TEA, Altas Habilidades ou Superdotação': 'PNE',
        'Pós Censo': 'Pós Censo',
        'Situação no Ano Selecionado': 'Situação',
        'Data do Último Procedimento': 'Data da situação'
    }
    # OBS: O código anterior do usuario (step 293) mudou 'Data de Matrícula' para 'Data Matrícula Suap na Unidade' no mapa_colunas??
    # Não, o diff mostrava 'Data de Matrícula': 'Data Matrícula Suap' e user alterou para 'Data Matrícula Suap'.
    # Vou usar KEYS que correspondem ao DF tratado em APP.PY ('Data de Matrícula').
    
    colunas_finais = list(mapa_colunas.values())
    
    # Renomear/Criar
    df_relatorio.rename(columns=mapa_colunas, inplace=True)
    
    # Garantir Inteiros na coluna '#'
    # if '#' in df_relatorio.columns: # Moved to inside group loop
    #     df_relatorio['#'] = df_relatorio['#'].astype(int)
        
    # Formatação de CPF (Máscara XXX.XXX.XXX-XX)
    if 'CPF' in df_relatorio.columns:
        def formatar_cpf_mascara(valor):
            if pd.isnull(valor) or str(valor).strip() == "": return ""
            nums = "".join(filter(str.isdigit, str(valor)))
            if len(nums) == 0: return str(valor)
            nums = nums.zfill(11) # Garante 11 digitos
            return f"{nums[:3]}.{nums[3:6]}.{nums[6:9]}-{nums[9:]}"
            
        df_relatorio['CPF'] = df_relatorio['CPF'].apply(formatar_cpf_mascara)

    # Identificar Grupos (Turmas)
    col_turma = "Turma no Ano Selecionado"
    grupos = []
    
    if col_turma in df_relatorio.columns:
        # Ordenar para garantir agrupamento visual se necessário
        # df_relatorio.sort_values(by=col_turma, inplace=True) # opcional, mas bom
        # Agrupar
        for nome_turma, df_group in df_relatorio.groupby(col_turma, sort=True):
             grupos.append(df_group.copy())
    else:
        grupos.append(df_relatorio.copy())
        
    # --- Loop Principal de Geração ---
    
    # Definições de Largura e Headers (fixos)
    larguras = {
        '#': 6, 
        'Grupo/Ano/Fase': 15,
        'Matrícula': 20,
        'CPF': 18,
        'Nome': 30,
        'Data de Nascimento': 15,
        'Idade (31/03)': 12,
        'Sexo': 8,
        'Filiação 1': 30,
        'Filiação 2': 30,
        'Naturalidade': 18,
        'Nacionalidade': 18,
        'Data Matrícula Suap': 15,
        'PNE': 8,
        'Pós Censo': 10, 
        'Situação': 18,
        'Data da situação': 13
    }
    # Ajuste manual para caber em 275mm (A4 Landscape usable)
    # Soma larguras atuais: 4+18+18+45+14+8+6+30+30+18+18+16+26+15+14 = 280.
    # Pode passar os 277mm (297 - 20 margem). Reduzir um pouco nomes/filiacao se precisar.
    
    larguras_lista = [larguras.get(col, 20) for col in colunas_finais]

    # Headers Visuais com Quebra de Linha
    headers_map = {
        'Grupo/Ano/Fase': 'Grupo/Ano/\nFase',
        'Idade (31/03)': 'Idade\n (31/03)',
        'Data da situação': 'Data da\n Situação',
        'Data de Nascimento': 'Data de\n Nascimento',
        'Data Matrícula Suap': 'Data \nMatrícula Suap',
        'Filiação 1': 'Filiação 1',
        'Filiação 1': 'Filiação 1',
        'Filiação 2': 'Filiação 2',
        'Pós Censo': 'Pós \n Censo'
    }
    # Obs: Coluna '#' não está no map, então usará o próprio nome '#', o que é correto.
    headers_texto = [headers_map.get(col, col) for col in colunas_finais]

    # Função Local para Print Row
    def print_row(pdf_instance, data_values, widths, is_header=False):
        line_height = 3 if is_header else 4
        # Altura mínima: 2 linhas de texto (4 * 2 = 8mm) se for dados
        min_h = 8 if not is_header else line_height
        
        start_y = pdf_instance.get_y()
        start_x = pdf_instance.get_x()
        
        # Colunas com alinhamento central
        center_cols = [
            '#',
            'Grupo/Ano/Fase', 
            'Matrícula', 
            'CPF', 
            'Data de Nascimento', 
            'Idade (31/03)', 
            'Sexo', 
            'Data Matrícula Suap', 
            'PNE', 
            'Pós Censo', 
            'Data da situação'
        ]
        
        # Passada 1: Calcular altura necessária (Line counting simulado)
        cell_info = []
        max_h = min_h
        
        for i, raw_val in enumerate(data_values):
            text = fix_text(raw_val)
            w = widths[i]
            col_name = colunas_finais[i]
            
            # Lógica de fonte reduzida
            current_font_size = 6
            special_font = False
            if not is_header:
                txt_str = str(text)
                if col_name == 'Nacionalidade' and 'Brasileira - Nascido no exterior' in txt_str:
                     current_font_size = 5
                     special_font = True
                elif col_name in ['Filiação 1', 'Filiação 2'] and 'NÃO CONSTA' in txt_str:
                     current_font_size = 5
                     special_font = True
            
            # Setar fonte para medição
            style = 'B' if is_header else ''
            pdf_instance.set_font('Arial', style, current_font_size)
            
            # Ajuste de largura útil para cálculo de quebra de linha
            c_margin = getattr(pdf_instance, 'c_margin', 1.0) # 1.0 default if missing, but usually ~0.35mm
            effective_w = w - (2 * c_margin)
            if effective_w < 0: effective_w = 0.1 # defensive
            
            # Contar linhas
            if not text:
                num_lines = 1
            else:
                num_lines = 0
                lines = text.split('\n')
                space_w = pdf_instance.get_string_width(' ')
                
                for line in lines:
                    if not line:
                        num_lines += 1
                        continue
                    
                    words = line.split(' ')
                    curr_line_w = 0
                    # Palavra por palavra
                    for idx_w, word in enumerate(words):
                        word_w = pdf_instance.get_string_width(word)
                        # Se primeira palavra (e cabe na largura ou é maior que largura total)
                        if curr_line_w == 0:
                            if word_w > effective_w:
                                # Palavra maior que a célula (vai quebrar?)
                                # FPDF normalmente imprime e estoura se for maior que w, mas MultiCell força wrap?
                                # Assumindo wrap conservador.
                                wrap_count = math.ceil(word_w / effective_w)
                                num_lines += wrap_count if wrap_count > 0 else 1
                                curr_line_w = 0 
                            else:
                                curr_line_w = word_w
                        else:
                            if curr_line_w + space_w + word_w > effective_w:
                                # Quebra linha
                                num_lines += 1
                                curr_line_w = word_w
                            else:
                                curr_line_w += space_w + word_w
                    num_lines += 1
            
            # Calcular altura desta celula
            cell_h = num_lines * line_height
            if cell_h > max_h:
                max_h = cell_h
            
            cell_info.append({
                'text': text,
                'w': w,
                'h': cell_h,
                'special_font': special_font,
                'font_size': current_font_size
            })

        # Restaurar fonte padrão
        pdf_instance.set_font('Arial', '', 6)

        # Checar Page Break Manual se a linha for muito alta (opcional, FPDF lida com multi_cell mas aqui desenhamos rect)
        # Se start_y estiver muito embaixo, add_page.
        page_h = 210
        margin_b = 10
        if start_y + max_h > page_h - margin_b:
             pdf_instance.add_page()
             if is_header:
                  # Se for header e quebrou página, reseta Y?
                  # print_table_header lida com isso se start_y for novo.
                  pass
             start_y = pdf_instance.get_y()
             start_x = pdf_instance.get_x()

        # Passada 2: Desenhar
        current_x = start_x
        
        for i, info in enumerate(cell_info):
            w = info['w']
            text = info['text']
            
            # Desenhar Borda/Fundo
            if is_header:
                pdf_instance.set_fill_color(220, 220, 220)
                pdf_instance.rect(current_x, start_y, w, max_h, 'FD')
            else:
                pdf_instance.rect(current_x, start_y, w, max_h, 'D')
            
            # alinhamento
            align = 'L'
            col_name = colunas_finais[i]
            if is_header:
                align = 'C'
            elif col_name in center_cols:
                align = 'C'
            
            # Fonte
            if info['special_font']:
                pdf_instance.set_font('Arial', '', info['font_size'])
            elif is_header:
                pdf_instance.set_font('Arial', 'B', 6)
            else:
                pdf_instance.set_font('Arial', '', 6)
            
            # Posição Y (Centralizada ou Topo)
            content_h = info['h']
            top_align_cols = ['Nome', 'Filiação 1', 'Filiação 2', 'Naturalidade', 'Nacionalidade', 'Situação']
            
            if not is_header and col_name in top_align_cols:
                y_offset = 0.5
            else:
                y_offset = (max_h - content_h) / 2
            
            # Renderizar Texto
            pdf_instance.set_xy(current_x, start_y + y_offset)
            pdf_instance.multi_cell(w, line_height, text, 0, align)
            
            current_x += w
            
        # Mover Y para fim da linha
        pdf_instance.set_y(start_y + max_h)
        # Resetar fonte
        pdf_instance.set_font('Arial', '', 6)

    # Função para desenhar o cabeçalho da tabela
    def print_table_header():
        pdf.set_font('Arial', 'B', 6)
        pdf.set_fill_color(220, 220, 220)
        print_row(pdf, headers_texto, larguras_lista, is_header=True)
        pdf.set_font('Arial', '', 6)

    # --- Iterar sobre Grupos ---
    for df_grupo in grupos:
        # Preparar Header Dinâmico para este grupo
        header_info = {}
        
        # 0. Turma
        if col_turma in df_grupo.columns:
            val = df_grupo[col_turma].iloc[0]
            header_info['turma'] = str(val) if pd.notna(val) else ""
        else:
            header_info['turma'] = "Única"
            
        # 1. Matriz
        # Tenta 'Matriz' primeiro, depois 'Curso' (que pode ser a matriz mapeada no DEPARA ou original)
        # Note: 'Matriz' pode ter sido renomeada para 'Matriz' se existia e mapeamos?
        # Verificando mapa_colunas: 'Curso': 'Matriz' foi comentado/tentado.
        # Vamos checar colunas disponíveis.
        matriz_val = ""
        if 'Matriz' in df_grupo.columns:
             matriz_val = df_grupo['Matriz'].iloc[0]
        elif 'Curso' in df_grupo.columns:
             matriz_val = df_grupo['Curso'].iloc[0]
        header_info['matriz'] = str(matriz_val) if pd.notna(matriz_val) else ""
             
        # 2. Data Base Censo (Fixo)
        header_info['data_censo'] = dados_escola.get('data_censo', '')
        
        # 3. Regra de Datas e Dias
        # Determinar contexto (Botão 1 ou 2)
        is_btn_2 = "EJA 2º SEM" in titulo_documento
        
        # Determinar Fase EJA
        desc_curso = ""
        if 'Descrição do Curso' in df_grupo.columns:
             desc_curso = str(df_grupo['Descrição do Curso'].iloc[0])
        
        # Normalizar string para comparação segura
        eja_phases = ['Educação de Jovens e Adultos Fases Finais', 'Educação de Jovens e Adultos Fases Iniciais']
        # Remover espaços extras se houver
        is_eja_phase = desc_curso.strip() in eja_phases
        
        enc_final = ""
        dias_final = ""
        
        # Regra 1 e 2
        if is_btn_2:
            # Botão 2 (EJA 2)
            # Regra: SE EJA Phase -> data_enc_eja2 / dias_eja2
            # Se não, fallback (assumindo também EJA 2 pois é botão EJA 2)
            enc_final = dados_escola.get('data_enc_eja2')
            dias_final = dados_escola.get('dias_eja2')
        else:
             # Botão 1 (Regular / EJA 1)
             if is_eja_phase:
                 enc_final = dados_escola.get('data_enc_eja1')
                 dias_final = dados_escola.get('dias_eja1')
             else:
                 enc_final = dados_escola.get('data_encerramento')
                 dias_final = dados_escola.get('total_dias_letivos')
                 
        header_info['data_encerramento'] = enc_final
        header_info['dias_letivos'] = dias_final
        # Armazenar Descrição do Curso para uso no Cabeçalho
        header_info['descricao_curso'] = desc_curso
        
        # Setar no objeto PDF
        pdf.current_header_info = header_info

        # Preparar DF do Grupo
        
        # --- Lógica de Deduplicação por CPF (Manter mais antiga) ---
        if 'CPF' in df_grupo.columns and 'Data Matrícula Suap' in df_grupo.columns: # Renomeada de Data de Matrícula
             try:
                 # Criar coluna temporária de data para ordenação
                 df_grupo['__data_sort'] = pd.to_datetime(df_grupo['Data Matrícula Suap'], dayfirst=True, errors='coerce')
                 # Ordenar por data (Ascendente = Mais Antiga Primeiro)
                 df_grupo.sort_values(by='__data_sort', ascending=True, inplace=True)
                 # Remover duplicatas de CPF, mantendo a primeira (mais antiga)
                 # Filtra apenas CPFs não nulos/vazios para evitar exclusão acidental de vazios distintos
                 mask_cpf_valid = df_grupo['CPF'].notna() & (df_grupo['CPF'] != "")
                 
                 # Separa dados com CPF e sem CPF
                 df_com_cpf = df_grupo[mask_cpf_valid].copy()
                 df_sem_cpf = df_grupo[~mask_cpf_valid].copy()
                 
                 # Remove dups nos com CPF
                 df_com_cpf.drop_duplicates(subset=['CPF'], keep='first', inplace=True)
                 
                 # Recombina (mantendo ordem de data se possível, ou reordenando por nome depois se desejado)
                 df_grupo = pd.concat([df_com_cpf, df_sem_cpf], ignore_index=True)
                 
             except Exception as e:
                 # Logar erro se necessário, ou pass silenciosamente mantendo dados
                 print(f"Erro na deduplicação: {e}")
                 pass
        
        # Ordenação Final: Primeiro por 'Ordenador', depois por 'Nome'
        cols_sort = []
        # Verifica se colunas existem antes de ordenar
        if 'Ordenador' in df_grupo.columns:
            # Garantir que Ordenador seja numérico ou string comparável consistentemente
            # Tentar converter para numérico se for o caso
            # df_grupo['Ordenador'] = pd.to_numeric(df_grupo['Ordenador'], errors='ignore') 
            cols_sort.append('Ordenador')
            
        if 'Nome' in df_grupo.columns:
            cols_sort.append('Nome')
            
        if cols_sort:
            df_grupo.sort_values(by=cols_sort, ascending=[True] * len(cols_sort), inplace=True)
            
        # 1. Resetar Sequencial '#'
        if '#' in df_grupo.columns: df_grupo.drop(columns=['#'], inplace=True)
        df_grupo.insert(0, '#', range(1, len(df_grupo) + 1))
        df_grupo['#'] = df_grupo['#'].astype(int)
        
        # 2. Preencher Colunas Faltantes
        for col_final in colunas_finais:
            if col_final not in df_grupo.columns:
                df_grupo[col_final] = ""
        
        # 3. Filtrar
        df_sub = df_grupo[colunas_finais]
        
        # 4. Paginação do Grupo
        registros_por_pagina = 15
        total_registros = len(df_sub)
        
        for i in range(0, total_registros, registros_por_pagina):
            chunk = df_sub.iloc[i : i + registros_por_pagina]
            
            pdf.add_page()
            
            # Header da Tabela
            print_table_header()
            
            # Rows
            for _, row in chunk.iterrows():
                row_values = [str(row[col]) for col in colunas_finais]
                print_row(pdf, row_values, larguras_lista, is_header=False)

    # Tratar retorno do fpdf que pode variar entre str e bytearray dependendo da versão/env
    val = pdf.output(dest='S')
    if isinstance(val, str):
        return val.encode('latin-1')
    return bytes(val)

def gerar_capa(dados_escola):
    """Gera a capa do Livro de Matrículas em PDF"""
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    
    # Configurações da Página
    page_width = 210
    page_height = 297
    margin_top = 10
    margin_left = 20
    margin_right = 10
    margin_bottom = 10
    
    # Configurar margens
    pdf.set_margins(margin_left, margin_top, margin_right)
    
    # 1. Borda de 2 pontos (aprox 0.7mm)
    pdf.set_line_width(0.7)
    pdf.rect(margin_left, margin_top, page_width - margin_left - margin_right, page_height - margin_top - margin_bottom)
    
    # Centralização e Alinhamento ao Alto
    y_cursor = margin_top + 15 # Espaço inicial após a borda
    
    # 2. Imagem Brasão
    # Centralizar imagem. Tamanho sugerido: 30x30mm
    logo_w = 30
    logo_h = 30
    try:
        # Centralizar na área útil (respeitando margens)
        # x_img = pdf.l_margin + (pdf.epw - logo_w) / 2
        # Ou manter centralizado na PÁGINA se for preferência? O texto usa cell(0) que respeita margens.
        # Vamos centralizar na PÁGINA para ficar simétrico visualmente no papel ou na área útil? 
        # Geralmente capa se centraliza no papel, mas com margem esquerda maior (encadernação), o centro visual muda.
        # Vou centralizar na área útil para alinhar com o texto.
        x_img = margin_left + (page_width - margin_left - margin_right - logo_w) / 2
        pdf.image('brasao.png', x=x_img, y=y_cursor, w=logo_w, h=logo_h)
        y_cursor += logo_h + 5
    except:
        # Se falhar imagem, apenas pula
        y_cursor += 10
        
    # 3. Textos Institucionais
    pdf.set_font('Arial', 'B', 14)
    original_y = pdf.get_y()
    pdf.set_y(y_cursor)
    
    textos_inst = [
        "Estado do Rio de Janeiro",
        "Prefeitura Municipal de Campos dos Goytacazes",
        "Secretaria Municipal de Educação, Ciência e Tecnologia"
        "Diretoria de Supervisão Escolar"
    ]
    
    for txt in textos_inst:
        pdf.cell(0, 8, txt, 0, 1, 'C')
        
    y_cursor = pdf.get_y() + 20
    
    # 4. Nome da Unidade
    nome_escola = fix_text(dados_escola.get('nome', 'Nome da Escola Não Informado'))
    pdf.set_y(y_cursor)
    pdf.set_font('Arial', 'B', 20)
    pdf.multi_cell(0, 12, nome_escola.upper(), 0, 'C')
    
    y_cursor = pdf.get_y() + 20
    
    # 5. Título com Ano Letivo
    ano = fix_text(dados_escola.get('ano_letivo', '____'))
    titulo = f"Livro de Matrículas - {ano}"
    
    pdf.set_y(y_cursor)
    pdf.set_font('Arial', 'B', 26)
    pdf.cell(0, 15, titulo, 0, 1, 'C')
    
    val = pdf.output(dest='S')
    if isinstance(val, str):
        return val.encode('latin-1')
    return bytes(val)

def gerar_termo_abertura(dados_escola):
    """Gera o Termo de Abertura em PDF"""
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    
    # Configurações da Página
    page_width = 210
    page_height = 297
    margin = 20 # Margem padrão A4 (Ref)
    
    # 1. Header (Adaptado da classe PDF para Portrait)
    # Configurar larguras e posições
    margin_left = 20 # Margem ajustada para encadernação (20mm)
    pdf.set_margins(margin_left, 10, 10) # Margem esquerda 20, direita 10
    
    # Ajustando para usar margem real da pagina definida no add_page que é default 1cm se nao setado? 
    # FPDF default margin is 10mm.
    # Mas no PDF class era set_margins(10,5,10).
    # Aqui vou usar margem 10mm lateral para o cabeçalho ficar largo.
    
    # Page width A4 Portrait
    page_width_p = 210
    usable_width = page_width_p - 2 * margin_left # 190mm
    
    # Dados Escola
    nome = fix_text(dados_escola.get('nome', ''))
    inep = fix_text(dados_escola.get('inep', ''))
    ano = fix_text(dados_escola.get('ano_letivo', ''))
    end = fix_text(dados_escola.get('logradouro', ''))
    num = fix_text(dados_escola.get('numero', ''))
    bairro = fix_text(dados_escola.get('bairro', ''))
    cep = fix_text(dados_escola.get('cep', ''))
    tel = fix_text(dados_escola.get('telefone', ''))
    email = fix_text(dados_escola.get('email', ''))

    pdf.set_y(5)  # Margem superior
    pdf.set_x(margin_left)
    
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    
    # --- Layout Colunas ---
    # Col 1: Logo (25mm)
    # Col 2: Info (Restante/2)
    # Col 3: Escola (Restante/2)
    
    w_col1 = 25
    remaining_w = usable_width - w_col1
    w_col2 = remaining_w * 0.45 # 45% do restante
    w_col3 = remaining_w * 0.55 # 55% para dados da escola que é mais longo
    
    # 1. Logo
    h_row1 = 20
    logo_size = 12.8
    try:
         # Alinhar logo à esquerda (x_start) para ficar igual à linha 2
         pdf.image('brasao.png', x_start, y_start + 2, logo_size, logo_size)
    except:
         pass
         
    # 2. Institucional (Ocupando o resto da linha 1)
    # w_col2 agora ocupa todo o espaço restante pois a escola desceu
    w_col2 = usable_width - w_col1
    
    pdf.set_xy(x_start + w_col1, y_start)
    pdf.set_font('Arial', '', 9)
    texto_inst = [
        "Estado do Rio de Janeiro",
        "Prefeitura Municipal de Campos dos Goytacazes",
        "Secretaria Municipal de Educação, Ciência e Tecnologia",
        "Diretoria de Supervisão Escolar"
    ]
    cur_x2 = x_start + w_col1 + 2
    cur_y2 = y_start + 2
    line_h2 = 4
    for linha in texto_inst:
        pdf.set_xy(cur_x2, cur_y2)
        pdf.cell(w_col2 - 4, line_h2, linha, 0, 0, 'L')
        cur_y2 += line_h2
        
    # Calcular altura da linha 1 para posicionar a linha 2
    y_end_row1 = max(y_start + h_row1, cur_y2 + 2)
    
    # 3. Dados Escola (Nova Linha 2)
    pdf.set_y(y_end_row1)
    pdf.set_x(margin_left)
    cur_y3 = pdf.get_y()
    
    # Formatar telefone
    tel_nums = "".join(filter(str.isdigit, str(tel)))
    if len(tel_nums) == 11:
        tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:7]}-{tel_nums[7:]}"
    elif len(tel_nums) == 10:
        tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:6]}-{tel_nums[6:]}"
    else:
        tel_fmt = tel
        
    linha1 = f"{nome}"
    linha2 = f"Endereço: {end} , {num}. {bairro}. CEP: {cep}"
    linha3 = f"Contatos: {tel_fmt} | {email}"
    linha4 = f"Código do INEP: {inep}"
    
    # Configurar Fonte 8 e Alinhamento Esquerda
    pdf.set_font('Arial', '', 8)
    line_h3 = 4
    
    # Imprimir linhas
    # Linha 1 (Nome) - Negrito opcional? O user não pediu negrito especificamente na instrução de "nova linha", 
    # mas antes era. Vou manter regular conforme pedido "linha 2 com fonte 8" (implica estilo base).
    # Se quiser Negrito no nome: pdf.set_font('Arial', 'B', 8)
    # Vou usar BOLD para o nome para destaque, e regular para o resto, ou tudo regular?
    # Pedido: "linha 2 com fonte 8". Vou por tudo regular para seguir estritamente, ou manter B para nome?
    # O user disse "conter os mesmos dados". 
    # Vou colocar Nome em Negrito 8 e o resto Regular 8 para ficar bonito.
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(usable_width, line_h3, linha1, 0, 1, 'L')
    
    pdf.set_font('Arial', '', 8)
    pdf.cell(usable_width, line_h3, linha2, 0, 1, 'L')
    pdf.cell(usable_width, line_h3, linha3, 0, 1, 'L')
    pdf.cell(usable_width, line_h3, linha4, 0, 1, 'L')
    
    cur_y3 = pdf.get_y()
    
    y_header = max(cur_y2, cur_y3) + 1
    pdf.line(margin_left, y_header, margin_left + usable_width, y_header)
    
    # Reset Y para baixo do cabeçalho
    # Ampliar a distância entre o cabeçalho e o TERMO DE ABERTURA em 5cm (50mm)
    pdf.set_y(y_start + h_row1 + 50)
    
    # --- 2. Título ---
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "TERMO DE ABERTURA DO LIVRO DE MATRÍCULAS", 0, 1, 'C')
    pdf.ln(10)
    
    # --- 3. Texto do Termo ---
    # Data Atual para o texto
    hoje = datetime.now()
    dia = hoje.day
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    mes = meses[hoje.month - 1]
    ano_atual = hoje.year
    
    texto_corpo = (
        f"Este livro contém ___________ folhas rubricadas por mim ___________________________________________, "
        f"matrícula _________________, servirá para registro de matrículas dos alunos da unidade escolar "
        f"{nome}, localizada no município de Campos dos Goytacazes (RJ)."
    )
    
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(0, 8, fix_text(texto_corpo), 0, 'J')
    
    pdf.ln(15)
    
    # Data Local
    txt_data = "Campos dos Goytacazes , ________ de ____________________ de _________."
    pdf.cell(0, 8, fix_text(txt_data), 0, 1, 'R')
    
    pdf.ln(30) # 3cm de espaçamento para as assinaturas
    
    # --- 4. Assinaturas ---
    # Coluna Width precisa ser re-calculada ou usar usable_width
    col_width = usable_width / 2

    y_sig = pdf.get_y()
    
    # Assinatura 1 (Esquerda) - Direção
    x_sig1 = margin_left
    pdf.set_xy(x_sig1, y_sig)
    
    # Desenhar linhas manualmente
    line_w = 60
    # Centro da coluna 1
    center_c1 = margin_left + (col_width / 2)
    pdf.line(center_c1 - line_w/2, y_sig, center_c1 + line_w/2, y_sig)
    
    pdf.set_xy(x_sig1, y_sig + 2)
    pdf.cell(col_width, 5, fix_text("Direção"), 0, 0, 'C')
    
    # Assinatura 2 (Direita) - Supervisor
    # Centro da coluna 2
    x_sig2 = margin_left + col_width
    center_c2 = x_sig2 + (col_width / 2)
    pdf.line(center_c2 - line_w/2, y_sig, center_c2 + line_w/2, y_sig)
    
    pdf.set_xy(x_sig1, y_sig + 2)
    pdf.cell(col_width, 5, fix_text("Direção"), 0, 0, 'C')
    
    # Assinatura 2 (Direita) - Supervisor
    # Centro da coluna 2
    x_sig2 = margin_left + col_width
    center_c2 = x_sig2 + (col_width / 2)
    pdf.line(center_c2 - line_w/2, y_sig, center_c2 + line_w/2, y_sig)
    
    pdf.set_xy(x_sig2, y_sig + 2)
    pdf.cell(col_width, 5, fix_text("Supervisor Pedagogo"), 0, 0, 'C')
    
    val = pdf.output(dest='S')
    if isinstance(val, str):
        return val.encode('latin-1')
    return bytes(val)

def gerar_termo_encerramento(dados_escola):
    """Gera o Termo de Encerramento em PDF"""
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    
    # Configurações da Página
    page_width = 210
    page_height = 297
    margin_left = 20 # Margem ajustada para encadernação (20mm)
    pdf.set_margins(margin_left, 10, 10) 
    
    # Page width A4 Portrait
    page_width_p = 210
    usable_width = page_width_p - 2 * margin_left
    
    # Dados Escola
    nome = fix_text(dados_escola.get('nome', ''))
    inep = fix_text(dados_escola.get('inep', ''))
    ano = fix_text(dados_escola.get('ano_letivo', ''))
    end = fix_text(dados_escola.get('logradouro', ''))
    num = fix_text(dados_escola.get('numero', ''))
    bairro = fix_text(dados_escola.get('bairro', ''))
    cep = fix_text(dados_escola.get('cep', ''))
    tel = fix_text(dados_escola.get('telefone', ''))
    email = fix_text(dados_escola.get('email', ''))

    pdf.set_y(5)  # Margem superior
    pdf.set_x(margin_left)
    
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    
    # --- Layout Colunas ---
    w_col1 = 25
    
    # 1. Logo
    h_row1 = 20
    logo_size = 12.8
    try:
         # Alinhar logo à esquerda (x_start)
         pdf.image('brasao.png', x_start, y_start + 2, logo_size, logo_size)
    except:
         pass
         
    # 2. Institucional (Ocupando o resto da linha 1)
    w_col2 = usable_width - w_col1
    
    pdf.set_xy(x_start + w_col1, y_start)
    pdf.set_font('Arial', '', 9)
    texto_inst = [
        "Estado do Rio de Janeiro",
        "Prefeitura Municipal de Campos dos Goytacazes",
        "Secretaria Municipal de Educação, Ciência e Tecnologia",
        "Diretoria de Supervisão Escolar"
    ]
    cur_x2 = x_start + w_col1 + 2
    cur_y2 = y_start + 2
    line_h2 = 4
    for linha in texto_inst:
        pdf.set_xy(cur_x2, cur_y2)
        pdf.cell(w_col2 - 4, line_h2, linha, 0, 0, 'L')
        cur_y2 += line_h2
        
    # Calcular altura da linha 1 para posicionar a linha 2
    y_end_row1 = max(y_start + h_row1, cur_y2 + 2)
    
    # 3. Dados Escola (Nova Linha 2)
    pdf.set_y(y_end_row1)
    pdf.set_x(margin_left)
    cur_y3 = pdf.get_y()
    
    # Formatar telefone
    tel_nums = "".join(filter(str.isdigit, str(tel)))
    if len(tel_nums) == 11:
        tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:7]}-{tel_nums[7:]}"
    elif len(tel_nums) == 10:
        tel_fmt = f"({tel_nums[:2]}) {tel_nums[2:6]}-{tel_nums[6:]}"
    else:
        tel_fmt = tel
        
    linha1 = f"{nome}"
    linha2 = f"Endereço: {end} , {num}. {bairro}. CEP: {cep}"
    linha3 = f"Contatos: {tel_fmt} | {email}"
    linha4 = f"Código do INEP: {inep}"
    
    # Configurar Fonte 8 e Alinhamento Esquerda
    pdf.set_font('Arial', '', 8)
    line_h3 = 4
    
    # Imprimir linhas
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(usable_width, line_h3, linha1, 0, 1, 'L')
    
    pdf.set_font('Arial', '', 8)
    pdf.cell(usable_width, line_h3, linha2, 0, 1, 'L')
    pdf.cell(usable_width, line_h3, linha3, 0, 1, 'L')
    pdf.cell(usable_width, line_h3, linha4, 0, 1, 'L')
    
    cur_y3 = pdf.get_y()
    
    y_header = max(cur_y2, cur_y3) + 1
    pdf.line(margin_left, y_header, margin_left + usable_width, y_header)
    
    # Reset Y para baixo do cabeçalho
    pdf.set_y(y_start + h_row1 + 50)
    
    # --- 2. Título ---
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "TERMO DE ENCERRAMENTO DO LIVRO DE MATRÍCULAS", 0, 1, 'C')
    pdf.ln(10)
    
    # --- 3. Texto do Termo ---
    texto_corpo = (
        f"Ficam encerradas na {nome}, para fins estatísticos as matrículas efetuadas "
        f"no ano de {ano} com um total de _______ alunos."
    )
    
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(0, 8, fix_text(texto_corpo), 0, 'J')
    
    pdf.ln(15)
    
    # Data Local
    txt_data = "Campos dos Goytacazes , ________ de ____________________ de _________."
    pdf.cell(0, 8, fix_text(txt_data), 0, 1, 'R')
    
    pdf.ln(30) # 3cm de espaçamento para as assinaturas
    
    # --- 4. Assinaturas ---
    # Coluna Width precisa ser re-calculada ou usar usable_width
    col_width = usable_width / 2

    y_sig = pdf.get_y()
    
    # Assinatura 1 (Esquerda) - Direção
    x_sig1 = margin_left
    pdf.set_xy(x_sig1, y_sig)
    
    # Desenhar linhas manualmente
    line_w = 60
    # Centro da coluna 1
    center_c1 = margin_left + (col_width / 2)
    pdf.line(center_c1 - line_w/2, y_sig, center_c1 + line_w/2, y_sig)
    
    pdf.set_xy(x_sig1, y_sig + 2)
    pdf.cell(col_width, 5, fix_text("Direção"), 0, 0, 'C')
    
    # Assinatura 2 (Direita) - Supervisor
    # Centro da coluna 2
    x_sig2 = margin_left + col_width
    center_c2 = x_sig2 + (col_width / 2)
    pdf.line(center_c2 - line_w/2, y_sig, center_c2 + line_w/2, y_sig)
    
    pdf.set_xy(x_sig2, y_sig + 2)
    pdf.cell(col_width, 5, fix_text("Supervisor Pedagogo"), 0, 0, 'C')
    
    val = pdf.output(dest='S')
    if isinstance(val, str):
        return val.encode('latin-1')
    return bytes(val)