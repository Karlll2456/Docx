"""create_document.py

Generates a serious DOCX legal opinion (parecer técnico) with:
- Uppercase title: 'PARECER TÉCNICO - CRIMES NAS CONTESTAÇÕES'
- Location/date header: 'Belém, 11 de dezembro de 2025'
- Default style: Arial 12
- Justified paragraphs
- Line spacing: 1.5
- Sections: EMENTA, RELATÓRIO, FUNDAMENTAÇÃO, CONCLUSÃO
- Signature line at the end

Output: parecer_tecnico.docx

Requires: python-docx
    pip install python-docx

Run:
    python create_document.py
"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


TITLE = "PARECER TÉCNICO - CRIMES NAS CONTESTAÇÕES"
HEADER_LOCAL_DATE = "Belém, 11 de dezembro de 2025"
OUTPUT_FILE = "parecer_tecnico.docx"
SIGNATURE = "José Ivanildo da Costa Navegantes Junior - OAB 23.953"


# NOTE:
# The user request references "the text from the earlier parecer content".
# Since this repository update is performed without access to that prior chat content,
# the section bodies below are intentionally written as placeholders.
# Replace the strings in SECTION_TEXT with the exact earlier parecer text.
SECTION_TEXT = {
    "EMENTA": (
        "(Substituir por ementa do parecer anterior.)\n"
        "Crimes nas contestações. Responsabilização penal por afirmações falsas em peças defensivas. "
        "Dever de veracidade. Limites da imunidade profissional."
    ),
    "RELATÓRIO": (
        "(Substituir pelo relatório do parecer anterior.)\n"
        "Cuida-se de consulta acerca da possibilidade de caracterização de ilícitos penais em razão de "
        "declarações lançadas em contestações judiciais, notadamente quando imputam fatos ou qualificações "
        "a terceiros sem suporte probatório mínimo, ou quando induzem o juízo a erro por meio de "
        "afirmações sabidamente inverídicas."
    ),
    "FUNDAMENTAÇÃO": (
        "(Substituir pela fundamentação do parecer anterior.)\n"
        "No âmbito penal, a veiculação de fatos falsos em peças processuais pode, conforme o caso concreto, "
        "ajustar-se a tipos como calúnia, difamação ou injúria, além de hipóteses envolvendo falsidade ideológica, "
        "fraude processual e denunciação caluniosa, observados os elementos objetivos e subjetivos de cada delito.\n"
        "A imunidade profissional do advogado não constitui salvo-conduto para práticas ilícitas; protege a "
        "manifestação técnica, nos limites da pertinência temática e da urbanidade, não alcançando a "
        "imputação dolosa de crime ou fato ofensivo dissociado do interesse de defesa.\n"
        "Deve-se considerar, ainda, o dever de lealdade processual e a vedação ao abuso do direito de "
        "defesa. A responsabilidade penal exige demonstração de dolo específico quando pertinente, bem como "
        "nexo de causalidade e adequada tipicidade."
    ),
    "CONCLUSÃO": (
        "(Substituir pela conclusão do parecer anterior.)\n"
        "Diante do exposto, conclui-se que a inserção de afirmações falsas e ofensivas em contestações pode, "
        "em tese, configurar crimes contra a honra e outros delitos correlatos, a depender do conteúdo, do contexto "
        "e da prova do elemento subjetivo. Recomenda-se redigir peças defensivas com estrita pertinência ao objeto "
        "litigioso, lastro fático mínimo e linguagem técnica, evitando imputações categóricas desacompanhadas de "
        "suporte probatório."
    ),
}


def _set_default_style(document: Document) -> None:
    """Set document default font to Arial 12."""
    style = document.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(12)


def _format_paragraph(paragraph) -> None:
    """Apply justification and 1.5 line spacing to a paragraph."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = paragraph.paragraph_format
    pf.line_spacing = 1.5


def add_heading(document: Document, text: str) -> None:
    p = document.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    _format_paragraph(p)


def add_body_paragraph(document: Document, text: str) -> None:
    # Support multi-paragraph strings separated by \n
    for piece in (text or "").split("\n"):
        piece = piece.strip()
        if not piece:
            continue
        p = document.add_paragraph(piece)
        _format_paragraph(p)


def build_document() -> Document:
    document = Document()
    _set_default_style(document)

    # Title
    title_p = document.add_paragraph()
    title_run = title_p.add_run(TITLE.upper())
    title_run.bold = True
    title_run.font.size = Pt(12)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.line_spacing = 1.5

    # Header/location/date
    header_p = document.add_paragraph(HEADER_LOCAL_DATE)
    header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_p.paragraph_format.line_spacing = 1.5

    document.add_paragraph()  # spacing line

    # Sections
    for section in ("EMENTA", "RELATÓRIO", "FUNDAMENTAÇÃO", "CONCLUSÃO"):
        add_heading(document, section)
        add_body_paragraph(document, SECTION_TEXT.get(section, ""))
        document.add_paragraph()  # spacing line

    # Signature line
    sig_p = document.add_paragraph(SIGNATURE)
    sig_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sig_p.paragraph_format.line_spacing = 1.5

    return document


def main() -> None:
    doc = build_document()
    doc.save(OUTPUT_FILE)


if __name__ == "__main__":
    main()
