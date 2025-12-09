from collections import defaultdict
import re
from docx import Document
from lxml import etree

# MCQ question start regex
QUESTION_START = re.compile(r"(?i)^(question(\s*\d+)?)\b|^\d+[\.\)]")

# Stop headers simplified: just check for 'essay type' or 'true or false' in the line
STOP_HEADERS_SIMPLE = re.compile(r"(?i)(essay\s*type|true\s*or\s*false)")

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _get_numbering_part_xml(doc):
    for part in doc.part.package.parts:
        if (
            part.content_type
            == "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
        ):
            xml_bytes = part.blob
            return etree.fromstring(xml_bytes)
    return None


def get_level_format_from_numId(doc, numId, ilvl):
    numId = str(numId)
    numbering_xml = _get_numbering_part_xml(doc)
    if numbering_xml is None:
        return None, None

    num_xpath = f".//w:num[@w:numId='{numId}']"
    num_elm = numbering_xml.find(num_xpath, namespaces=W_NS)
    if num_elm is None:
        return None, None

    abs_node = num_elm.find("./w:abstractNumId", namespaces=W_NS)
    if abs_node is None:
        return None, None
    abstract_num_id = abs_node.get(f"{{{W_NS['w']}}}val")

    abs_xpath = f".//w:abstractNum[@w:abstractNumId='{abstract_num_id}']"
    abs_elm = numbering_xml.find(abs_xpath, namespaces=W_NS)
    if abs_elm is None:
        return None, None

    lvl_xpath = f"./w:lvl[@w:ilvl='{ilvl}']"
    lvl_elm = abs_elm.find(lvl_xpath, namespaces=W_NS)
    if lvl_elm is None:
        return None, None

    numFmt_elm = lvl_elm.find("./w:numFmt", namespaces=W_NS)
    lvlText_elm = lvl_elm.find("./w:lvlText", namespaces=W_NS)

    numFmt = numFmt_elm.get(f"{{{W_NS['w']}}}val") if numFmt_elm is not None else None
    lvlText = (
        lvlText_elm.get(f"{{{W_NS['w']}}}val") if lvlText_elm is not None else None
    )

    return numFmt, lvlText


def convert_level_to_number(n, fmt):
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

        # Stop if line contains essay type or true/false
        if STOP_HEADERS_SIMPLE.search(ln):
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
    qs = []
    i = 0
    qnum = 1

    while i < len(lines):
        ln = lines[i]

        # Stop parsing if essay or true/false section starts
        if STOP_HEADERS_SIMPLE.search(ln):
            break

        if QUESTION_START.search(ln):
            question_text = ln
            j = i + 1
            q_lines = []
            while (
                j < len(lines)
                and not QUESTION_START.search(lines[j])
                and not STOP_HEADERS_SIMPLE.search(lines[j])
            ):
                if lines[j].strip():
                    q_lines.append(lines[j].strip())
                j += 1

            options = q_lines[:4]
            options = options + [""] * (4 - len(options))
            qs.append({"number": qnum, "question": question_text, "options": options})
            qnum += 1
            i = j
        else:
            i += 1

    return qs


def create_output_doc(questions, answers, out_path, remaining_text):
    doc = Document()

    for q in questions:
        qnum = q["number"]
        qtext = q["question"]

        doc.add_paragraph(f"Question {qnum}) {qtext}")

        opts = q["options"]
        letters = ["a", "b", "c", "d", "e"]

        for letter, opt in zip(letters, opts):
            doc.add_paragraph(f"{letter}. {opt}")

        if qnum - 1 < len(answers):
            doc.add_paragraph(f"Answer: {answers[qnum - 1]}")
        doc.add_paragraph("")

    # Add remaining text (essay, true/false, etc.)
    if remaining_text:
        doc.add_page_break()
        for ln in remaining_text:
            doc.add_paragraph(ln)

    doc.save(out_path)


def process_docx(path, output_path):
    doc = Document(path)
    lines = []
    counters = defaultdict(lambda: defaultdict(int))
    last_level_for_num = defaultdict(lambda: -1)

    for p in doc.paragraphs:
        p_elm = p._p
        numPr = None
        if p_elm.pPr is not None and p_elm.pPr.numPr is not None:
            numPr = p_elm.pPr.numPr

        if numPr is None:
            lines.append(p.text)
            continue

        numId = numPr.numId.val
        ilvl = int(numPr.ilvl.val) if numPr.ilvl is not None else 0

        if last_level_for_num[numId] == -1 or ilvl <= last_level_for_num[numId]:
            for deeper in list(counters[numId].keys()):
                if deeper > ilvl:
                    del counters[numId][deeper]

        counters[numId][ilvl] += 1
        last_level_for_num[numId] = ilvl

        numFmt, lvlText = get_level_format_from_numId(doc, numId, ilvl)

        if numFmt == "bullet" or (
            lvlText is not None and "%" not in lvlText and numFmt == "bullet"
        ):
            prefix = (lvlText or "•") + " "
        else:
            val = convert_level_to_number(counters[numId][ilvl], numFmt)
            if lvlText is None:
                prefix = val + ". "
            else:
                placeholder = f"%{ilvl + 1}"
                prefix = lvlText.replace(placeholder, val) + " "

        lines.append(prefix + p.text)

    # Remaining text: everything after first line containing essay type or true/false
    remaining_index = None
    for i, ln in enumerate(lines):
        if STOP_HEADERS_SIMPLE.search(ln):
            remaining_index = i
            break
    remaining_text = lines[remaining_index:] if remaining_index is not None else []

    # Parse only MCQs before essay/true-false
    parse_until = remaining_index if remaining_index is not None else len(lines)
    answers = parse_answer_key(lines[:parse_until])
    questions = parse_questions(lines[:parse_until])

    create_output_doc(questions, answers, output_path, remaining_text)
    print("Document processed and saved.")


if __name__ == "__main__":
    input_file = "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new.docx"
    output_file = (
        "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new_processed_new.docx"
    )
    process_docx(input_file, output_file)
