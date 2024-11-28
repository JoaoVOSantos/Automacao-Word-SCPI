import tkinter as tk
from tkinter import filedialog
from docx import Document
from docx.shared import Pt
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
    # Alterar nome da primeira coluna para "Lances"
    if tabela.rows:
        # Verificar se a tabela tem pelo menos uma linha (cabeçalho)
        cabeçalho = tabela.rows[0]
        if len(cabeçalho.cells) > 0:
            # Alterar o texto da célula no índice 0 para "Lances"
            cabeçalho.cells[0].text = "Lances"

    for linha in tabela.rows:
        for idx, celula in enumerate(linha.cells):
            # Se a célula for da coluna de índice 1 (que contém os nomes), aplicar a capitalização
            if idx == 1:
                for par in celula.paragraphs:
                    for run in par.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(10)
                    # Modificar o texto da célula, capitalizando corretamente os nomes
                    celula.text = capitalizar_nome(celula.text)

            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                par.alignment = 1  # Centraliza o texto horizontalmente
                par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                par.paragraph_format.line_spacing = 1.15
                par.paragraph_format.space_before = Pt(6)

            # Adicionar bordas nas células da tabela 2
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



# Função para aplicar estilos às tabelas
def estilizar_tabelas(doc):
    for i, tabela in enumerate(doc.tables):
        if i == 0:  # Primeira tabela
            estilizar_tabela1(tabela)
        elif i == 1:  # Segunda tabela
            estilizar_tabela2(tabela)
    return doc


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