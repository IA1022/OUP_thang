from collections import defaultdict
import re
from docx import Document
from lxml import etree

QUESTION_RE = re.compile(r"(?i)^question\b")
OPTION_RE = re.compile(r"(?i)^[a-e][\.\)]\s*(.*)")  # a. text  OR  a) text
ANSWER_RE = re.compile(r"(?i)^answer\s*key")
ANSWER_ITEM_RE = re.compile(r"(?i)^(\d+)[\.\)]?\s*([a-e])$|^([a-e])$")

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _get_numbering_part_xml(doc):
    """
    Return an lxml element for numbering.xml (the numbering part).
    """
    # find the part with numbering content type
    for part in doc.part.package.parts:
        if (
            part.content_type
            == "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
        ):
            xml_bytes = part.blob  # bytes of numbering.xml
            return etree.fromstring(xml_bytes)
    return None


def get_level_format_from_numId(doc, numId, ilvl):
    """
    Given a Document, numId (string or int) and ilvl (int),
    return (numFmt, lvlText) or (None, None) if not found.

    numFmt is like 'decimal', 'lowerLetter', 'bullet', ...
    lvlText is like '%1.', '%2)', '•', etc.
    """
    numId = str(numId)
    numbering_xml = _get_numbering_part_xml(doc)
    if numbering_xml is None:
        return None, None

    # 1) find the <w:num w:numId="..."> element
    num_xpath = f".//w:num[@w:numId='{numId}']"
    num_elm = numbering_xml.find(num_xpath, namespaces=W_NS)
    if num_elm is None:
        return None, None

    # 2) get abstractNumId (value of <w:abstractNumId w:val="..."/>)
    abs_node = num_elm.find("./w:abstractNumId", namespaces=W_NS)
    if abs_node is None:
        return None, None
    abstract_num_id = abs_node.get(f"{{{W_NS['w']}}}val")

    # 3) find the <w:abstractNum w:abstractNumId="..."> element
    abs_xpath = f".//w:abstractNum[@w:abstractNumId='{abstract_num_id}']"
    abs_elm = numbering_xml.find(abs_xpath, namespaces=W_NS)
    if abs_elm is None:
        return None, None

    # 4) find the <w:lvl w:ilvl="..."> element inside abstractNum
    lvl_xpath = f"./w:lvl[@w:ilvl='{ilvl}']"
    lvl_elm = abs_elm.find(lvl_xpath, namespaces=W_NS)
    if lvl_elm is None:
        return None, None

    # 5) numFmt and lvlText children
    numFmt_elm = lvl_elm.find("./w:numFmt", namespaces=W_NS)
    lvlText_elm = lvl_elm.find("./w:lvlText", namespaces=W_NS)

    numFmt = numFmt_elm.get(f"{{{W_NS['w']}}}val") if numFmt_elm is not None else None
    lvlText = (
        lvlText_elm.get(f"{{{W_NS['w']}}}val") if lvlText_elm is not None else None
    )

    return numFmt, lvlText


def convert_level_to_number(n, fmt):
    """Convert 1 → 1, a, i depending on numFormat"""
    if fmt == "decimal":
        return str(n)
    if fmt == "lowerLetter":
        return chr(ord("a") + (n - 1))
    if fmt == "upperLetter":
        return chr(ord("A") + (n - 1))
    if fmt == "lowerRoman":
        return to_roman(n).lower()
    if fmt == "upperRoman":
        return to_roman(n)
    return str(n)


def to_roman(n):
    vals = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ]
    res = ""
    for v, r in vals:
        while n >= v:
            res += r
            n -= v
    return res


def clean_option(line):
    """Remove leading a., a) etc."""
    m = OPTION_RE.match(line.strip())
    return m.group(1).strip() if m else line.strip()


def parse_answer_key(lines):
    answers = []
    for ln in lines:
        ln = ln.strip().lower()
        m = ANSWER_ITEM_RE.match(ln)
        if not m:
            continue
        # Match formats:
        # 1. a   => m.group(2)
        # a     => m.group(3)
        ans = m.group(2) or m.group(3)
        answers.append(ans)
    return answers


# extract answers with answer number and questions and options
def process_docx(path, output_path):
    doc = Document(path)
    lines = []
    counters = defaultdict(lambda: defaultdict(int))
    last_level_for_num = defaultdict(lambda: -1)
    for p in doc.paragraphs:
        p_elm = p._p  # lxml-ish object from python-docx (oxml)
        numPr = None
        if p_elm.pPr is not None and p_elm.pPr.numPr is not None:
            numPr = p_elm.pPr.numPr

        if numPr is None:
            # non-numbered paragraph
            lines.append(p.text)
            # reset last levels? not necessary
            continue

        numId = numPr.numId.val
        ilvl = int(numPr.ilvl.val) if numPr.ilvl is not None else 0

        # reset counters when jumping to a new higher-level list or when ilvl < last_level
        if last_level_for_num[numId] == -1 or ilvl <= last_level_for_num[numId]:
            # when moving to same-or-higher (smaller ilvl number) level, zero-out deeper levels
            for deeper in list(counters[numId].keys()):
                if deeper > ilvl:
                    del counters[numId][deeper]

        # increment current level counter
        counters[numId][ilvl] += 1
        last_level_for_num[numId] = ilvl

        numFmt, lvlText = get_level_format_from_numId(doc, numId, ilvl)

        if numFmt == "bullet" or (
            lvlText is not None and "%" not in lvlText and numFmt == "bullet"
        ):
            # bullet: lvlText typically contains the bullet glyph
            prefix = (lvlText or "•") + " "
        else:
            # compute actual number/letter/roman
            val = convert_level_to_number(counters[numId][ilvl], numFmt)
            if lvlText is None:
                # fallback: just use decimal
                prefix = val + ". "
            else:
                # lvlText contains placeholders like "%1.", "%2)"
                # replace placeholder corresponding to this level number (levels are 0-based but placeholders are 1-based)
                placeholder = f"%{ilvl + 1}"
                prefix = lvlText.replace(placeholder, val) + " "

        lines.append(prefix + p.text)

    # Split into two sections: questions + answer key
    try:
        ak_index = next(i for i, ln in enumerate(lines) if ANSWER_RE.search(ln))
    except StopIteration:
        raise ValueError("Answer key not found")

    question_lines = lines[:ak_index]
    answer_lines = lines[ak_index + 1 :]

    # Parse answer key
    answers = parse_answer_key(answer_lines)

    # Parse questions
    formatted_questions = []
    current_q = None
    current_opts = []

    for ln in question_lines:
        if not ln.strip():
            continue

        # New question starts
        if QUESTION_RE.search(ln):
            if current_q:
                # store previous question
                formatted_questions.append((current_q, current_opts))
            current_q = ln
            current_opts = []
            continue

        # Option line (a., b., c., etc.)
        if OPTION_RE.match(ln):
            current_opts.append(clean_option(ln))
            continue

        # Option line without labels
        if current_q:
            current_opts.append(ln.strip())

    # Store last block
    if current_q:
        formatted_questions.append((current_q, current_opts))

    # Build output doc
    out = Document()

    for i, (qtext, opts) in enumerate(formatted_questions, start=1):
        # Clean question text: remove "Question ..."
        qtext_clean = re.sub(r"(?i)^question\b[\s:\-\d\.]*", "", qtext).strip()

        out.add_paragraph(f"Question {i}) {qtext_clean}")

        # Only take first 4 or 5 options
        opts = opts[:5]

        labels = ["a", "b", "c", "d", "e"]
        for label, opt in zip(labels, opts):
            out.add_paragraph(f"{label}. {opt}")

        # Add answer
        if i <= len(answers):
            out.add_paragraph(f"Answer: {answers[i - 1]}")
        out.add_paragraph("")

    out.save(output_path)


if __name__ == "__main__":
    input_file = "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new.docx"
    output_file = (
        "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new_processed4.docx"
    )
    process_docx(input_file, output_file)
