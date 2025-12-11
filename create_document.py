"""create_document.py

Gera o arquivo `parecer_tecnico.docx` com formatação padrão para Parecer Técnico.

Requisitos atendidos:
- TITLE em maiúsculas (saída) com o valor: "PARECER TÉCNICO - CRIMES NAS CONTESTAÇÕES".
- Cabeçalho: "Belém, 11 de dezembro de 2025".
- Assinatura: "José Ivanildo da Costa Navegantes Junior - OAB 23.953".
- Fonte Arial 12.
- Parágrafos justificados.
- Espaçamento 1,5.
- Preenche a seção SECTION_TEXT com os textos de EMENTA, RELATÓRIO, FUNDAMENTAÇÃO, CONCLUSÃO fornecidos pelo usuário.
- Preserva parágrafos: separa por linhas em branco e também quebra itens de lista no início da linha
  em parágrafos separados.

Dependência:
    pip install python-docx

Uso sugerido:
    python create_document.py --ementa "..." --relatorio "..." --fundamentacao "..." --conclusao "..."

Ou via STDIN (para textos grandes):
    python create_document.py --stdin
"""

from __future__ import annotations

import argparse
import re
from typing import Iterable, List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt


TITLE = "PARECER TÉCNICO - CRIMES NAS CONTESTAÇÕES"
HEADER_TEXT = "Belém, 11 de dezembro de 2025"
SIGNATURE_TEXT = "José Ivanildo da Costa Navegantes Junior - OAB 23.953"
OUTPUT_DOCX = "parecer_tecnico.docx"


_LIST_ITEM_RE = re.compile(
    r"^\s*(?:"
    r"(?:[-*•])|"  # bullets
    r"(?:\d+[\.)])|"  # 1. / 1)
    r"(?:[a-zA-Z][\.)])|"  # a. / a)
    r"(?:[ivxlcdmIVXLCDM]+[\.)])"  # i. / IV)
    r")\s+(.+?)\s*$"
)


def _normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def split_into_paragraphs(text: str) -> List[str]:
    """Divide o texto em parágrafos.

    Regras:
    - Parágrafos são separados por uma ou mais linhas em branco.
    - Linhas iniciadas com marcadores/numeração (itens de lista) viram parágrafos separados.
    - Mantém a ordem e remove espaços laterais.
    """

    text = _normalize_newlines(text).strip("\n")
    if not text.strip():
        return []

    blocks = re.split(r"\n\s*\n+", text)

    paragraphs: List[str] = []
    for block in blocks:
        lines = [ln.rstrip() for ln in block.split("\n") if ln.strip()]
        if not lines:
            continue

        # Se houver itens de lista, cada linha item vira um parágrafo.
        # Caso contrário, junta as linhas do bloco em um único parágrafo.
        any_list_item = any(_LIST_ITEM_RE.match(ln) for ln in lines)
        if any_list_item:
            for ln in lines:
                m = _LIST_ITEM_RE.match(ln)
                if m:
                    paragraphs.append(m.group(0).strip())
                else:
                    # Linha normal dentro do bloco (ex.: texto introdutório antes da lista)
                    paragraphs.append(ln.strip())
        else:
            paragraphs.append(" ".join(ln.strip() for ln in lines).strip())

    return [p for p in paragraphs if p]


def _set_run_font(run) -> None:
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(12)


def add_formatted_paragraph(document: Document, text: str) -> None:
    """Adiciona um parágrafo justificado, Arial 12, espaçamento 1,5."""

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    pf = p.paragraph_format
    pf.line_spacing = 1.5

    run = p.add_run(text)
    _set_run_font(run)


def add_heading_like(document: Document, text: str) -> None:
    """Adiciona um título/seção em destaque simples (negrito) mantendo formatação base."""

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    pf = p.paragraph_format
    pf.line_spacing = 1.5

    run = p.add_run(text)
    run.bold = True
    _set_run_font(run)


def build_section_text(ementa: str, relatorio: str, fundamentacao: str, conclusao: str) -> str:
    """Monta o conteúdo completo (SECTION_TEXT) a partir das seções fornecidas."""

    parts = []
    if ementa.strip():
        parts.append("EMENTA")
        parts.append(ementa.strip())
    if relatorio.strip():
        parts.append("RELATÓRIO")
        parts.append(relatorio.strip())
    if fundamentacao.strip():
        parts.append("FUNDAMENTAÇÃO")
        parts.append(fundamentacao.strip())
    if conclusao.strip():
        parts.append("CONCLUSÃO")
        parts.append(conclusao.strip())

    return "\n\n".join(parts).strip()


def generate_docx(ementa: str, relatorio: str, fundamentacao: str, conclusao: str) -> str:
    doc = Document()

    # Título em maiúsculas (saída)
    add_heading_like(doc, TITLE.upper())

    # Cabeçalho (linha simples, justificado)
    add_formatted_paragraph(doc, HEADER_TEXT)

    # Corpo (SECTION_TEXT)
    section_text = build_section_text(ementa, relatorio, fundamentacao, conclusao)

    for para in split_into_paragraphs(section_text):
        # Para rótulos de seção (EMENTA, RELATÓRIO, ...), destaca como heading-like
        if para.strip().upper() in {"EMENTA", "RELATÓRIO", "FUNDAMENTAÇÃO", "CONCLUSÃO"}:
            add_heading_like(doc, para.strip().upper())
        else:
            add_formatted_paragraph(doc, para)

    # Assinatura
    add_formatted_paragraph(doc, "")
    add_formatted_paragraph(doc, SIGNATURE_TEXT)

    doc.save(OUTPUT_DOCX)
    return OUTPUT_DOCX


def _read_stdin_sections() -> tuple[str, str, str, str]:
    """Lê um texto do STDIN contendo seções rotuladas.

    Formato esperado (flexível):
        EMENTA
        ...

        RELATÓRIO
        ...

        FUNDAMENTAÇÃO
        ...

        CONCLUSÃO
        ...

    Se algum rótulo não for encontrado, o conteúdo correspondente fica vazio.
    """

    import sys

    raw = sys.stdin.read()
    raw = _normalize_newlines(raw)

    labels = ["EMENTA", "RELATÓRIO", "FUNDAMENTAÇÃO", "CONCLUSÃO"]

    # Captura blocos por rótulos em linhas próprias
    pattern = re.compile(
        r"(?ms)^\s*(EMENTA|RELATÓRIO|FUNDAMENTAÇÃO|CONCLUSÃO)\s*$\n(.*?)(?=^\s*(?:EMENTA|RELATÓRIO|FUNDAMENTAÇÃO|CONCLUSÃO)\s*$|\Z)"
    )

    found = {k: "" for k in labels}
    for m in pattern.finditer(raw):
        found[m.group(1)] = m.group(2).strip("\n")

    return (
        found["EMENTA"],
        found["RELATÓRIO"],
        found["FUNDAMENTAÇÃO"],
        found["CONCLUSÃO"],
    )


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Gera parecer_tecnico.docx (Arial 12, justificado, 1.5)")
    parser.add_argument("--ementa", default="", help="Texto da EMENTA")
    parser.add_argument("--relatorio", default="", help="Texto do RELATÓRIO")
    parser.add_argument("--fundamentacao", default="", help="Texto da FUNDAMENTAÇÃO")
    parser.add_argument("--conclusao", default="", help="Texto da CONCLUSÃO")
    parser.add_argument(
        "--stdin",
        action="store_true",
        help="Lê do STDIN um documento com seções rotuladas (EMENTA/RELATÓRIO/FUNDAMENTAÇÃO/CONCLUSÃO)",
    )

    args = parser.parse_args(list(argv) if argv is not None else None)

    if args.stdin:
        ementa, relatorio, fundamentacao, conclusao = _read_stdin_sections()
    else:
        ementa, relatorio, fundamentacao, conclusao = (
            args.ementa,
            args.relatorio,
            args.fundamentacao,
            args.conclusao,
        )

    generate_docx(ementa, relatorio, fundamentacao, conclusao)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
