import tkinter as tk
from tkinter import filedialog
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.enum.text import WD_LINE_SPACING
import os

                
# Função para estilizar os parágrafos
def estilizar_paragrafos(doc):
    for par in doc.paragraphs:
        for run in par.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(10)
        par.paragraph_format.left_indent = Pt(0)
        par.paragraph_format.right_indent = Pt(0)
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        par.paragraph_format.line_spacing = 1.15

# Função para definir margens e cabeçalhos
def definir_margens(doc):
    for section in doc.sections:
        section.top_margin = Pt(2.23 * 28.35)
        section.bottom_margin = Pt(1.5 * 28.35)
        section.left_margin = Pt(2 * 28.35)
        section.right_margin = Pt(1.5 * 28.35)
        section.header_distance = Pt(0.5 * 28.35)
        section.footer_distance = Pt(0.56 * 28.35)

# Função para formatar os nomes com a primeira letra maiúscula
def formatar_nomes(celula):
    for par in celula.paragraphs:
        # Formatar texto: primeira letra maiúscula
        par.text = " ".join([palavra.capitalize() for palavra in par.text.split()])

# Função para estilizar a Tabela 1
def estilizar_tabela1(tabela):
    # Detectar o índice da coluna "Nome"
    nome_coluna_index = None
    for idx, celula in enumerate(tabela.rows[0].cells):
        if "Nome" in celula.text:
            nome_coluna_index = idx
            break

    # Ajustar as larguras das colunas conforme especificado
    larguras = {
        0: 1.51,  # Portaria
        1: 2.00,  # Data
        2: 5.25,  # Nome
        3: 3.75,  # Cargo
        4: 2.75,  # CPF
        5: 2.26   # RG
    }

    for i, linha in enumerate(tabela.rows):
        for idx, celula in enumerate(linha.cells):
            # Definir largura das colunas
            if idx in larguras:
                largura = larguras[idx]
                celula._element.get_or_add_tcPr().append(
                    parse_xml(
                        f'<w:tcW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="{int(largura * 567)}" w:type="dxa"/>' 
                    )
                )

            # Centralizar o texto horizontalmente
            for par in celula.paragraphs:
                par.alignment = 1  # Centralizado
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)
            celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Verificar e formatar se for a coluna "Nome"
            if idx == nome_coluna_index:
                formatar_nomes(celula)

            # Aplicar bordas em todas as células
            tc_pr = celula._element.get_or_add_tcPr()
            borders = parse_xml(
                r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'</w:tcBorders>' 
            )
            tc_pr.append(borders)

    # Estilizar a tabela inteira após a alteração dos nomes
    for linha in tabela.rows:
        for celula in linha.cells:
            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                par.alignment = 1
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)

def capitalizar_nome(nome):
    # Lista de exceções
    excecoes = ["LTDA", "EPP", "S.A.", "S.A", "SA", "ME"]
    
    # Quebra o nome em palavras e capitaliza cada uma, exceto as exceções
    palavras = nome.split()
    palavras_capitalizadas = []
    
    for palavra in palavras:
        if palavra.upper() in excecoes:
            palavras_capitalizadas.append(palavra.upper())  # Mantém a exceção em maiúsculo
        else:
            palavras_capitalizadas.append(palavra.capitalize())  # Capitaliza a palavra normalmente
    
    # Junta as palavras novamente em uma string
    return " ".join(palavras_capitalizadas)


def estilizar_tabela2(tabela):
    # Alterar nome da primeira coluna para "Lances" (apenas uma vez)
    if tabela.rows:
        cabeçalho = tabela.rows[0]
        if len(cabeçalho.cells) > 0:
            cabeçalho.cells[0].text = "Lances"
    
    # Definir as larguras das colunas em centímetros
    larguras_colunas = [1.26, 7.00, 3.00, 3.50, 2.76]  # Larguras em cm

    # Definir larguras das colunas em pontos
    larguras_colunas_pts = [largura * 28.35 for largura in larguras_colunas]

    for linha in tabela.rows:
        # Definir a largura das colunas
        for col_idx, celula in enumerate(linha.cells):
            if col_idx < len(larguras_colunas_pts):
                largura_em_pontos = int(larguras_colunas_pts[col_idx])
                celula._element.get_or_add_tcPr().append(parse_xml(
                    f'<w:tcW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="{largura_em_pontos}" w:type="dxa"/>'))

    for linha in tabela.rows:
        for idx, celula in enumerate(linha.cells):
            # Se a célula for da coluna de índice 0 (Lances), remover espaços vazios
            if idx == 0:
                for par in celula.paragraphs:
                    par.text = ' '.join(par.text.split())  # Remove espaços extras

            # Se a célula for da coluna de índice 1 (que contém os nomes), aplicar a capitalização
            if idx == 1:
                # Preservar quebras de linha
                novo_texto = []
                for par in celula.paragraphs:
                    # Capitaliza cada parágrafo, preservando a quebra de linha
                    par.text = capitalizar_nome(par.text)
                    novo_texto.append(par.text)
                # Recria o texto da célula com as quebras de linha
                celula.text = '\n'.join(novo_texto)

            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                par.alignment = 1  # Centraliza o texto horizontalmente
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)

            # Adicionar bordas nas células
            tc_pr = celula._element.get_or_add_tcPr()
            borders = parse_xml(
                r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'</w:tcBorders>' 
            )
            tc_pr.append(borders)

            # Centralizar verticalmente a célula
            celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Estilizar a tabela inteira após a alteração das bordas
    for linha in tabela.rows:
        for celula in linha.cells:
            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                par.alignment = 1  # Centraliza o texto horizontalmente
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)


def ajustar_largura_por_tabela(tabela):
    """
    Ajusta as larguras das colunas com base na tabela.
    A largura das colunas será configurada conforme o tipo de tabela.
    """
    # Definir larguras para a tabela com as colunas: Item, Lote, Descrição, Valor Total, Status Lance
    larguras_tabela1 = [1.21, 1.8, 10.5, 2, 2.01]

    # Verifica se a tabela tem as colunas específicas e aplica as larguras
    if len(tabela.columns) >= 5:  # Verifica se há pelo menos 5 colunas
        # Verifica os títulos das colunas para identificar a estrutura
        colunas_identificadas = [tabela.rows[0].cells[i].text.strip() for i in range(5)]
        
        # Se a estrutura for igual à esperada, define as larguras
        if colunas_identificadas == ['Item', 'Lote', 'Descrição', 'Valor Total', 'Status Lance']:
            for idx, largura in enumerate(larguras_tabela1):
                if idx < len(tabela.columns):
                    for celula in tabela.columns[idx].cells:
                        celula.width = Cm(largura)

def capitalizar_nome_com_quebra_linha(nome):
    # Lista de exceções
    excecoes = ["LTDA", "EPP", "S.A.", "S.A", "SA", "ME"]
    
    # Quebra o nome em palavras, preservando as quebras de linha
    partes = nome.split("\n")
    partes_capitalizadas = []

    for parte in partes:
        palavras = parte.split()
        palavras_capitalizadas = []

        for palavra in palavras:
            # Verifica se a palavra é uma exceção
            if palavra.upper() in excecoes:
                palavras_capitalizadas.append(palavra.upper())  # Mantém a exceção em maiúsculo
            else:
                palavras_capitalizadas.append(palavra.capitalize())  # Capitaliza a palavra normalmente

        # Junta as palavras capitalizadas novamente
        partes_capitalizadas.append(" ".join(palavras_capitalizadas))
    
    # Junta novamente as partes com quebra de linha
    return "\n".join(partes_capitalizadas)

def estilizar_tabela3(tabela):
    # Determinar as colunas a remover ou estilizar
    col_lote_idx = None
    col_codigo_idx = None
    col_valor_unitario_idx = None
    col_unidade_idx = None
    col_quantidade_idx = None
    col_marca_idx = None
    col_descricao_idx = None  # Para armazenar o índice da coluna "Descrição"

    # Verifica a primeira linha como cabeçalho para identificar as colunas
    if tabela.rows:
        for idx, celula in enumerate(tabela.rows[0].cells):
            texto_celula = celula.text.strip()
            if texto_celula.startswith("Lote"):
                col_lote_idx = idx
            elif texto_celula.startswith("Código"):
                col_codigo_idx = idx
            elif texto_celula.startswith("Valor Unitário"):
                col_valor_unitario_idx = idx
            elif texto_celula.startswith("Unidade"):
                col_unidade_idx = idx
            elif texto_celula.startswith("Quantidade"):
                col_quantidade_idx = idx
            elif texto_celula.startswith("Marca"):
                col_marca_idx = idx
            elif texto_celula.startswith("Descrição"):
                col_descricao_idx = idx  # Identificar o índice da coluna "Descrição"

    # Ajusta o índice da coluna "Lote" se estiver misturado com "Código"
    if col_codigo_idx is not None and col_lote_idx is None:
        for linha in tabela.rows:
            texto_celula = linha.cells[col_codigo_idx].text.strip()
            if "Lote" in texto_celula.split("\n")[0]:
                col_lote_idx = col_codigo_idx
                col_codigo_idx = None
                break

    # Se existir uma coluna "Código", removê-la
    if col_codigo_idx is not None:
        for linha in tabela.rows:
            linha.cells[col_codigo_idx]._element.getparent().remove(linha.cells[col_codigo_idx]._element)

    # Recalcular os índices das colunas após a remoção de "Código"
    col_lote_idx, col_valor_unitario_idx, col_unidade_idx, col_quantidade_idx, col_marca_idx, col_descricao_idx = None, None, None, None, None, None
    if tabela.rows:
        for idx, celula in enumerate(tabela.rows[0].cells):
            texto_celula = celula.text.strip()
            if texto_celula.startswith("Lote"):
                col_lote_idx = idx
            elif texto_celula.startswith("Valor Unitário"):
                col_valor_unitario_idx = idx
            elif texto_celula.startswith("Unidade"):
                col_unidade_idx = idx
            elif texto_celula.startswith("Quantidade"):
                col_quantidade_idx = idx
            elif texto_celula.startswith("Marca"):
                col_marca_idx = idx
            elif texto_celula.startswith("Descrição"):
                col_descricao_idx = idx  # Identificar novamente o índice da coluna "Descrição"

    # Ajusta larguras com base nas colunas da tabela
    ajustar_largura_por_tabela(tabela)

    # Estilizar a tabela restante
    for linha in tabela.rows:
        for idx, celula in enumerate(linha.cells):
            # Aplicar bordas em todas as células
            tc_pr = celula._element.get_or_add_tcPr()
            borders = parse_xml(
                r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>' 
                r'</w:tcBorders>' 
            )
            tc_pr.append(borders)

            # Centralizar verticalmente a célula
            celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Estilizar todas as células restantes
    for linha in tabela.rows:
        for idx, celula in enumerate(linha.cells):
            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.name = 'Arial'  # Garantir que a fonte seja Arial
                    run.font.size = Pt(9)
                par.alignment = 1  # Centralizar texto horizontalmente
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)

            # Aplicar a função capitalizar_nome_com_quebra_linha na coluna "Descrição"
            if idx == col_descricao_idx:
                # Para cada célula na coluna "Descrição", capitalizar o texto preservando a quebra de linha
                celula.text = capitalizar_nome_com_quebra_linha(celula.text)

                # Após a capitalização, aplicar novamente a estilização na célula
                for par in celula.paragraphs:
                    for run in par.runs:
                        run.font.name = 'Arial'  # Garantir que a fonte seja Arial
                        run.font.size = Pt(10)  # Manter o tamanho da fonte
                    par.alignment = 1  # Centralizar texto horizontalmente
                    par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                    par.paragraph_format.line_spacing = 1.15
                    par.paragraph_format.space_before = Pt(6)

                


# Função para aplicar estilos às tabelas
def estilizar_tabelas(doc):
    for i, tabela in enumerate(doc.tables):
        if i == 0:
            # Estilizar a primeira tabela com estilo de Tabela 1
            estilizar_tabela1(tabela)
        elif i == 1:
            # Estilizar a segunda tabela com estilo de Tabela 2
            estilizar_tabela2(tabela)
        elif i == 2:
            # Estilizar a terceira tabela com estilo de Tabela 3
            estilizar_tabela3(tabela)
        else:
            # Para outras tabelas, você pode aplicar um estilo genérico ou personalizado
            ajustar_largura_por_tabela(tabela)  # Ajustar larguras de tabela
            for linha in tabela.rows:
                for celula in linha.cells:
                    for par in celula.paragraphs:
                        for run in par.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
                        par.alignment = 1  # Centraliza o texto horizontalmente
                        par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                        par.paragraph_format.line_spacing = 1.15
                        par.paragraph_format.space_before = Pt(6)
                    # Adicionar bordas em todas as células
                    tc_pr = celula._element.get_or_add_tcPr()
                    borders = parse_xml(
                        r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        r'</w:tcBorders>'
                    )
                    tc_pr.append(borders)

                    # Centralizar verticalmente a célula
                    celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


# Função para copiar conteúdo e estilizar
def copiar_conteudo_para_modelo():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione um arquivo Word", filetypes=[("Documentos Word", "*.docx")])
    if caminho_arquivo:
        doc_origem = Document(caminho_arquivo)
        caminho_modelo = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Modelo.docx')
        
        if os.path.exists(caminho_modelo):
            doc_modelo = Document(caminho_modelo)
            for elemento in doc_origem.element.body:
                doc_modelo.element.body.append(elemento)
            
            estilizar_paragrafos(doc_modelo)
            definir_margens(doc_modelo)
            estilizar_tabelas(doc_modelo)

            caminho_saida = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documentos Word", "*.docx")])
            if caminho_saida:
                doc_modelo.save(caminho_saida)
                print(f"Arquivo salvo em: {caminho_saida}")
        else:
            print(f"Modelo não encontrado: {caminho_modelo}")

# Interface gráfica
def criar_interface():
    root = tk.Tk()
    root.withdraw()
    copiar_conteudo_para_modelo()

# Executa o programa
criar_interface() 