import re
from docx import Document

MCQ_HEADER_REGEX = r"\d+\.\t"
ANSWER_PATTERN = r"(\d+)\s*[\.\:\-\)]?\s*([a-eA-E])"


def get_num_info(paragraph):
    """
    Returns:
        (is_numbered, numId, ilvl)
    or (False, None, None)
    """
    p = paragraph._p
    if p.pPr is None or p.pPr.numPr is None:
        return False, None, None

    numPr = p.pPr.numPr
    numId = numPr.numId.val
    ilvl = numPr.ilvl.val
    return True, numId, ilvl


def get_level_format(doc, numId, ilvl):
    """Get format and level text like '%1.' or '%2)' or '•' """
    numbering_def = doc.part.numbering_part.numbering_definitions.get(numId)
    if numbering_def is None:
        return None, None

    lvl = numbering_def.levels[ilvl]
    return lvl.numFormat, lvl.level_text


def convert_level_to_number(n, fmt):
    """Convert 1 → 1, a, i depending on numFormat"""
    if fmt == 'decimal':
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

# extract answers with answer number and questions and options
def process_docx(path, output_path):
    doc = Document(path)
    full_text = ""
    for p in doc.paragraphs:
        is_numbered, numId, ilvl = get_num_info(p)
        
        if not is_numbered:
            full_text += p.text
            continue

        # Initialize counter structure
        if numId not in counters:
            counters[numId] = {}

        # Reset sublevel counters when needed
        for level in range(0, ilvl + 1):
            counters[numId].setdefault(level, 0)

        # Increment only this level
        counters[numId][ilvl] += 1

        numFormat, lvlText = get_level_format(doc, numId, ilvl)

         # Bullet
        if numFormat == 'bullet':
            prefix = lvlText + " "
        else:
            # Compute actual value (1, a, i…)
            val = convert_level_to_number(counters[numId][ilvl], numFormat)

            # Replace the placeholder %1, %2, etc.
            prefix = lvlText.replace(f"%{ilvl+1}", val) + " "

        full_text += (prefix + p.text)
    
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
    input_file = (
        "C:/Users/agarwais/Downloads/DocFiles/9780199039678/Albanese_5e_TB_ch1.docx"
    )
    output_file = "C:/Users/agarwais/Downloads/Variation2_processed.docx"
    process_docx(input_file, output_file)
