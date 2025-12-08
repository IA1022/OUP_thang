from collections import defaultdict
import re
from docx import Document
from lxml import etree

MCQ_HEADER_REGEX = r"\d+\.\t"
ANSWER_PATTERN = r"(\d+)\s*[\.\:\-\)]?\s*([a-eA-E])"

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


# ---------- Detect manual numbering ----------
MANUAL_PATTERN = re.compile(
    r"^([0-9]+|[a-zA-Z]|[ivxlcdmIVXLCDM]+)[\.\)\:]\s+"
)

def detect_manual_numbering(text):
    """
    Returns (prefix, remaining_text) if manual numbering is found.
    Example: "3) Hello" -> ("3)", "Hello")
    """
    m = MANUAL_PATTERN.match(text.strip())
    if not m:
        return None, text
    prefix = m.group(0).strip()
    remaining = text[len(prefix):].strip()
    return prefix, remaining


# ---------- Roman util ----------
def to_roman(n):
    vals = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]
    res = ""
    for v, r in vals:
        while n >= v:
            res += r
            n -= v
    return res


# ---------- Convert Word numbering format ----------
def convert_level_to_number(n, fmt):
    if fmt in (None, 'decimal'):
        return str(n)
    if fmt == 'lowerLetter':
        return chr(ord('a') + (n - 1))
    if fmt == 'upperLetter':
        return chr(ord('A') + (n - 1))
    if fmt == 'lowerRoman':
        return to_roman(n).lower()
    if fmt == 'upperRoman':
        return to_roman(n)
    return str(n)


# ---------- Read numbering.xml ----------
def get_numbering_xml(doc):
    for part in doc.part.package.parts:
        if part.content_type == \
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
            return etree.fromstring(part.blob)
    return None

def get_level_format(doc, numId, ilvl):
    numId = str(numId)
    xml = get_numbering_xml(doc)
    if xml is None:
        return None, None

    num = xml.find(f".//w:num[@w:numId='{numId}']", namespaces=W_NS)
    if num is None:
        return None, None

    abstract_id = num.find("./w:abstractNumId", namespaces=W_NS)
    if abstract_id is None:
        return None, None
    abs_num_id = abstract_id.get(f"{{{W_NS['w']}}}val")

    abstract = xml.find(
        f".//w:abstractNum[@w:abstractNumId='{abs_num_id}']",
        namespaces=W_NS
    )
    if abstract is None:
        return None, None

    lvl = abstract.find(f"./w:lvl[@w:ilvl='{ilvl}']", namespaces=W_NS)
    if lvl is None:
        return None, None

    numFmt = lvl.find("./w:numFmt", namespaces=W_NS)
    lvlText = lvl.find("./w:lvlText", namespaces=W_NS)

    return (
        numFmt.get(f"{{{W_NS['w']}}}val") if numFmt is not None else None,
        lvlText.get(f"{{{W_NS['w']}}}val") if lvlText is not None else None,
    )


# extract answers with answer number and questions and options
def process_docx(path, output_path):
    doc = Document(path)
    lines = []
    counters = defaultdict(lambda: defaultdict(int))
    last_level_for_num = defaultdict(lambda: -1)
    for p in doc.paragraphs:
        text = p.text.strip()
        # 1️⃣ Check for Word auto-numbering
        numPr = p._p.pPr.numPr if p._p.pPr is not None else None
        if numPr is not None:
            numId = numPr.numId.val
            ilvl = int(numPr.ilvl.val) if numPr.ilvl is not None else 0

            numFmt, lvlText = get_level_format(doc, numId, ilvl)

            # Reset deeper levels
            for deeper in list(counters[numId]):
                if deeper > ilvl:
                    del counters[numId][deeper]

            counters[numId][ilvl] += 1
            last_level[numId] = ilvl

            # Create prefix
            if numFmt == "bullet":
                prefix = lvlText + " "
            else:
                n = counters[numId][ilvl]
                real = convert_level_to_number(n, numFmt)
                prefix = lvlText.replace(f"%{ilvl+1}", real) + " "

            lines.append(prefix + text)
            continue
        
        # 2️⃣ Else, check if user manually typed numbering
        manual = detect_manual_numbering(text)
        if manual[0]:
            prefix, rest = manual
            print(prefix + " " + rest)
            continue

        # 3️⃣ Just normal text
        lines.append(text)

    full_text = "\n".join(lines)
    match = re.search(r"Answer Key", full_text, flags=re.I)
    if not match:
        print("No 'Answer Key:' section found.")
        return
    else:
        print("Answer key found.")

    # Split text into before and after answer key
    answer_key_start = match.start()
    before = full_text[:answer_key_start]
    after = full_text[answer_key_start:]

    answers = dict(re.findall(ANSWER_PATTERN, after))
    answers = {int(k): v.lower() for k, v in answers.items()}

    # If numbering is missing, assign incremental numbers
    if not answers:
        raw_answers = re.findall(r"\b[a-eA-E]\b", after)
        answers = {i + 1: raw_answers[i].lower() for i in range(len(raw_answers))}

    # Split questions
    question_blocks = re.split(r"\n(?=\d+\.\t)", before)
    new_doc = Document()

    for block in question_blocks:
        text = block.strip()

        if not text:
            continue

        if re.search(MCQ_HEADER_REGEX, text):
            lines = text.split("\n")
            qnum_match = re.match(r"(\d+)", lines[0])

            if qnum_match:
                qnum = int(qnum_match.group(1))

                # Extract question text and options
                question_text = "\n".join(lines[:5])
                options = {
                    "a": lines[-5] if len(lines) >= 5 else "",
                    "b": lines[-4] if len(lines) >= 4 else "",
                    "c": lines[-3] if len(lines) >= 3 else "",
                    "d": lines[-2] if len(lines) >= 2 else "",
                    "e": lines[-1] if len(lines) >= 1 else "",
                }

                # add question, options and answers to new document
                new_doc.add_paragraph(f"Q{qnum}: {question_text}")
                for opt in ["a", "b", "c", "d", "e"]:
                    new_doc.add_paragraph(f"{opt}) {options[opt]}")

                if qnum in answers:
                    new_doc.add_paragraph(f"Answer: {answers[qnum]}")
                    new_doc.add_paragraph("")
                else:
                    new_doc.add_paragraph(text)

            new_doc.save(output_path)
            print("Document processed and saved.")


if __name__ == "__main__":
    input_file = "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new.docx"
    output_file = (
        "C:\\Users\\nikhi\\OneDrive\\Desktop\\Scripts\\test_new_processed.docx"
    )
    process_docx(input_file, output_file)
