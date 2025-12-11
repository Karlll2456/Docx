"""
Script para criar um documento DOCX de exemplo
Autor: Gerado automaticamente
Data: 2025-12-11
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def criar_documento():
    """
    Cria um documento DOCX com conteúdo de exemplo
    """
    try:
        # Criar um novo documento
        doc = Document()
        
        # Adicionar título
        titulo = doc.add_heading('Documento de Exemplo', 0)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Adicionar parágrafo de introdução
        paragrafo1 = doc.add_paragraph(
            'Este é um documento de exemplo criado usando Python e a biblioteca python-docx. '
            'Esta biblioteca permite criar e modificar documentos do Microsoft Word de forma programática.'
        )
        
        # Adicionar parágrafo com formatação
        paragrafo2 = doc.add_paragraph('Este parágrafo contém ')
        paragrafo2.add_run('texto em negrito').bold = True
        paragrafo2.add_run(', ')
        paragrafo2.add_run('texto em itálico').italic = True
        paragrafo2.add_run(' e ')
        run_colorido = paragrafo2.add_run('texto colorido')
        run_colorido.font.color.rgb = RGBColor(255, 0, 0)
        paragrafo2.add_run('.')
        
        # Adicionar subtítulo
        doc.add_heading('Lista de Recursos', level=1)
        
        # Adicionar lista não ordenada
        doc.add_paragraph('Criação de parágrafos', style='List Bullet')
        doc.add_paragraph('Formatação de texto (negrito, itálico, cores)', style='List Bullet')
        doc.add_paragraph('Adição de títulos e subtítulos', style='List Bullet')
        doc.add_paragraph('Criação de listas', style='List Bullet')
        doc.add_paragraph('Inserção de tabelas', style='List Bullet')
        
        # Adicionar subtítulo para tabela
        doc.add_heading('Exemplo de Tabela', level=1)
        
        # Adicionar tabela
        tabela = doc.add_table(rows=4, cols=3)
        tabela.style = 'Light Grid Accent 1'
        
        # Cabeçalho da tabela
        celulas_cabecalho = tabela.rows[0].cells
        celulas_cabecalho[0].text = 'Nome'
        celulas_cabecalho[1].text = 'Idade'
        celulas_cabecalho[2].text = 'Cidade'
        
        # Dados da tabela
        dados = [
            ('João Silva', '25', 'São Paulo'),
            ('Maria Santos', '30', 'Rio de Janeiro'),
            ('Pedro Oliveira', '28', 'Belo Horizonte')
        ]
        
        for i, (nome, idade, cidade) in enumerate(dados, start=1):
            celulas = tabela.rows[i].cells
            celulas[0].text = nome
            celulas[1].text = idade
            celulas[2].text = cidade
        
        # Adicionar conclusão
        doc.add_heading('Conclusão', level=1)
        doc.add_paragraph(
            'Este documento demonstra as principais funcionalidades da biblioteca python-docx. '
            'Você pode expandir este código para criar documentos mais complexos conforme suas necessidades.'
        )
        
        # Salvar o documento
        nome_arquivo = 'exemplo.docx'
        doc.save(nome_arquivo)
        print(f'✓ Documento "{nome_arquivo}" criado com sucesso!')
        
        return nome_arquivo
        
    except Exception as e:
        print(f'✗ Erro ao criar o documento: {str(e)}')
        return None

if __name__ == '__main__':
    print('Criando documento DOCX...')
    criar_documento()
