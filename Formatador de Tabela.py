import tkinter as tk
from tkinter import filedialog
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml  # Corrigido: Importação do parse_xml
from docx.enum.text import WD_LINE_SPACING

# Função para aplicar o estilo nas tabelas
def estilizar_tabelas(doc):
    for tabela in doc.tables:
        # Aplica borda em todas as células
        for linha in tabela.rows:
            for celula in linha.cells:
                # Centralizar o texto horizontalmente
                for par in celula.paragraphs:
                    par.alignment = 1  # 1 = centralizado
                    # Definir fonte Arial e tamanho 10 para cada parágrafo
                    for run in par.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(10)
                        
                        
                    # Adicionar espaçamento entre as linhas (1,15)
                    par = celula.paragraphs[0]  # Exemplo de como acessar um parágrafo na célula
                    par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE  # Define como múltiplo
                    par.paragraph_format.line_spacing = 1.15  # Define o fator de multiplicação
        
                    # Adicionar espaçamento antes da linha (6pt)
                    par.paragraph_format.space_before = Pt(6)
                
                # Centralizar verticalmente
                celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Aplicar borda
                for borda in celula._element.xpath('.//w:tcPr//w:tcBorders'):
                    borda.clear()  # Limpar bordas existentes
                celula._element.get_or_add_tcPr().append(
                    parse_xml(  # Agora usando parse_xml corretamente
                        r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                        r'<w:top w:val="single" w:space="0" w:size="4" w:color="000000"/>'
                        r'<w:left w:val="single" w:space="0" w:size="4" w:color="000000"/>'
                        r'<w:bottom w:val="single" w:space="0" w:size="4" w:color="000000"/>'
                        r'<w:right w:val="single" w:space="0" w:size="4" w:color="000000"/>'
                        r'</w:tcBorders>'
                    )
                )
    return doc

# Função para abrir o arquivo Word
def abrir_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo Word", filetypes=[("Documentos Word", "*.docx")]
    )
    if caminho_arquivo:
        doc = Document(caminho_arquivo)
        doc = estilizar_tabelas(doc)
        # Salvar o arquivo estilizado com novo nome
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho_saida:
            doc.save(caminho_saida)
            print(f"Arquivo salvo em: {caminho_saida}")
        else:
            print("Erro ao salvar o arquivo.")

# Criando a interface gráfica
root = tk.Tk()
root.withdraw()  # Esconde a janela principal

# Abre a janela para selecionar o arquivo
abrir_arquivo()
