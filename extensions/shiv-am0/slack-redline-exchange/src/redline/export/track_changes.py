from __future__ import annotations

import io
import re
from datetime import UTC, datetime

import docx
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement

from ..core.store import ProposalRow

SIDE_AUTHORS = {"vendor": "Vendor", "customer": "Customer"}


def _clean_html(html_text: str | None) -> str:
    import html as html_mod

    if not html_text:
        return ""
    return html_mod.unescape(re.sub(r"<[^>]+>", "", html_text)).strip()


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def _run(text: str, tag: str = "w:r") -> OxmlElement:
    r = OxmlElement("w:r")
    text_tag = "w:delText" if tag == "w:delText" else "w:t"
    t = OxmlElement(text_tag)
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _del_run(text: str) -> OxmlElement:
    r = OxmlElement("w:r")
    d = OxmlElement("w:delText")
    d.set(qn("xml:space"), "preserve")
    d.text = text
    r.append(d)
    return r


def _wrap(kind: str, run: OxmlElement, author: str, date: str, rev_id: int) -> OxmlElement:
    wrapper = OxmlElement(kind)
    wrapper.set(qn("w:id"), str(rev_id))
    wrapper.set(qn("w:author"), author)
    wrapper.set(qn("w:date"), date)
    wrapper.append(run)
    return wrapper


class _RevisionIdCounter:
    def __init__(self) -> None:
        self._value = 0

    def next(self) -> int:
        self._value += 1
        return self._value


def _find_paragraph(document: docx.Document, needle: str):
    target = _normalize(needle)
    if not target:
        return None
    for paragraph in document.paragraphs:
        if _normalize(paragraph.text) == target:
            return paragraph
    for paragraph in document.paragraphs:
        if target in _normalize(paragraph.text):
            return paragraph
    return None


def _apply_edit(
    document: docx.Document, proposal: ProposalRow, author: str, date: str, ids: _RevisionIdCounter
) -> bool:
    old_text = _clean_html(proposal.old_html)
    paragraph = _find_paragraph(document, old_text)
    if paragraph is None:
        return False
    p = paragraph._p
    for run_element in p.findall(qn("w:r")):
        p.remove(run_element)
    new_text = _clean_html(proposal.new_html)
    p.append(_wrap("w:del", _del_run(old_text), author, date, ids.next()))
    if new_text:
        p.append(_wrap("w:ins", _run(new_text), author, date, ids.next()))
    return True


def _apply_delete(
    document: docx.Document, proposal: ProposalRow, author: str, date: str, ids: _RevisionIdCounter
) -> bool:
    old_text = _clean_html(proposal.old_html)
    paragraph = _find_paragraph(document, old_text)
    if paragraph is None:
        return False
    p = paragraph._p
    for run_element in p.findall(qn("w:r")):
        p.remove(run_element)
    p.append(_wrap("w:del", _del_run(old_text), author, date, ids.next()))
    return True


def _apply_create(
    document: docx.Document, proposal: ProposalRow, author: str, date: str, ids: _RevisionIdCounter
) -> None:
    new_text = _clean_html(proposal.new_html)
    paragraph = document.add_paragraph()
    p = paragraph._p
    p.append(_wrap("w:ins", _run(new_text), author, date, ids.next()))


def _append_heading(document: docx.Document, text: str, level: int = 1) -> None:
    document.add_heading(text, level=level)


def _append_table(document: docx.Document, rows: list[list[str]]) -> None:
    if not rows:
        return
    table = document.add_table(rows=1, cols=len(rows[0]))
    style_names = {s.name for s in document.styles}
    table.style = "Light Grid Accent 1" if "Light Grid Accent 1" in style_names else None
    header_cells = table.rows[0].cells
    headers = ["Change", "Proposed by", "Vendor", "Customer", "Feedback"]
    for i, h in enumerate(headers):
        header_cells[i].text = h
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = value


def _proposal_summary(proposal: ProposalRow) -> str:
    old = _clean_html(proposal.old_html)
    new = _clean_html(proposal.new_html)
    if proposal.operation == "delete":
        return f"Delete: {old[:120]}"
    if proposal.operation == "create":
        return f"Add: {new[:120]}"
    return f"Edit: {old[:60]} -> {new[:60]}"


def build_redline_docx(
    original_filename: str, original_bytes: bytes, proposals: list[ProposalRow]
) -> bytes:
    """Build a lawyer-facing docx with OOXML track changes, generated locally from the
    original source document plus the recorded negotiation trail. Committed changes are
    rendered inline as w:ins/w:del attributed to the proposing side; rejected proposals are
    listed honestly in an appendix rather than silently dropped; changes whose original text
    could not be confidently located are called out explicitly rather than guessed at.
    """
    if original_filename.lower().endswith(".docx"):
        document = docx.Document(io.BytesIO(original_bytes))
    else:
        # Non-docx originals (md/txt/html) have no OOXML structure to overlay changes onto.
        # We still produce a usable docx, but it carries the negotiation trail as an
        # appendix rather than inline w:ins/w:del marks on the original prose.
        document = docx.Document()
        document.add_heading("Source document (converted, no inline redlines)", level=1)
        for line in original_bytes.decode("utf-8", errors="replace").splitlines():
            if line.strip():
                document.add_paragraph(line.strip())

    date = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    ids = _RevisionIdCounter()

    inline_capable = original_filename.lower().endswith(".docx")
    committed = [p for p in proposals if p.state == "committed"]
    rejected = [p for p in proposals if p.state == "rejected"]
    unmatched: list[ProposalRow] = []

    if inline_capable:
        for proposal in committed:
            author = SIDE_AUTHORS.get(proposal.proposed_by_side, proposal.proposed_by_side)
            if proposal.operation == "edit":
                ok = _apply_edit(document, proposal, author, date, ids)
            elif proposal.operation == "delete":
                ok = _apply_delete(document, proposal, author, date, ids)
            else:
                _apply_create(document, proposal, author, date, ids)
                ok = True
            if not ok:
                unmatched.append(proposal)
    else:
        unmatched.extend(committed)

    if unmatched:
        _append_heading(document, "Unmatched Changes", level=1)
        document.add_paragraph(
            "The following approved changes could not be confidently located in the "
            "original document text and are listed here instead of being guessed at:"
        )
        for proposal in unmatched:
            document.add_paragraph(_proposal_summary(proposal), style=None)

    if rejected:
        _append_heading(document, "Appendix: Rejected Proposals", level=1)
        rows = [
            [
                _proposal_summary(p),
                SIDE_AUTHORS.get(p.proposed_by_side, p.proposed_by_side),
                p.vendor_decision or "-",
                p.customer_decision or "-",
                (p.vendor_feedback or p.customer_feedback or "-"),
            ]
            for p in rejected
        ]
        try:
            _append_table(document, rows)
        except Exception:
            for row in rows:
                document.add_paragraph(" | ".join(row))

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


__all__ = ["build_redline_docx"]
