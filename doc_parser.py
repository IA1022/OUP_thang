from collections import defaultdict
import re
from docx import Document
from lxml import etree

QUESTION_START = re.compile(r"(?i)^(question(\s*\d+)?)\b|^\d+[\.\)]")

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


def parse_answer_key(lines):
    """
    Extract answer key.
    Supports formats:
      A
      C
      D
    or:
      1. a
      2. c
    """
    answers = []
    start = None

    for i, ln in enumerate(lines):
        if re.search(r"(?i)^answer\s*key", ln):
            start = i + 1
            break

    if start is None:
        return answers

    for ln in lines[start:]:
        ln = ln.strip()
        if not ln:
            continue

        # Stop if new section begins
        if re.search(r"(?i)^(type|chapter)", ln):
            break

        # Format 1: single letter per line
        if re.fullmatch(r"[A-Ea-e]", ln):
            answers.append(ln.lower())
            continue

        # Format 2: "1. a"
        m = re.match(r"(\d+)[\.\)]\s*([a-eA-E])", ln)
        if m:
            answers.append(m.group(2).lower())

    return answers


def parse_questions(lines):
    """Parse MCQ questions + options."""
    qs = []
    i = 0
    qnum = 1

    while i < len(lines):
        ln = lines[i]

        # detect question start
        if QUESTION_START.search(ln):
            question_text = ln

            # collect question body (same paragraph)
            j = i + 1
            q_lines = []
            while (
                j < len(lines)
                and not QUESTION_START.search(lines[j])
                and not re.search(r"(?i)^answer\s*key", lines[j])
            ):
                if lines[j].strip():
                    q_lines.append(lines[j].strip())
                j += 1

            # question text = first line
            # options = next 4 lines
            options = q_lines[:4]
            options = options + [""] * (4 - len(options))  # pad if needed

            qs.append({"number": qnum, "question": question_text, "options": options})
            qnum += 1

            i = j
        else:
            i += 1

    return qs


def create_output_doc(questions, answers, out_path):
    doc = Document()

    for q in questions:
        qnum = q["number"]
        qtext = q["question"]

        doc.add_paragraph(f"Question {qnum}) {qtext}")

        opts = q["options"]
        letters = ["a", "b", "c", "d"]

        for letter, opt in zip(letters, opts):
            doc.add_paragraph(f"{letter}. {opt}")

        if qnum - 1 < len(answers):
            doc.add_paragraph(f"Answer: {answers[qnum - 1]}")
        doc.add_paragraph("")

    doc.save(out_path)


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

    # full_text = "\n".join(lines)
    # input(full_text)
    answers = parse_answer_key(lines)
    questions = parse_questions(lines)

    create_output_doc(questions, answers, output_file)
    print("Document processed and saved.")


if __name__ == "__main__":
    input_file = "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new.docx"
    output_file = (
        "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new_processed2.docx"
    )
    process_docx(input_file, output_file)
