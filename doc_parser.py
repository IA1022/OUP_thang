import re
from docx import Document

MCQ_HEADER_REGEX = r"\d+\.\t"
ANSWER_PATTERN = r"(\d+)\s*[\.\:\-\)]?\s*([a-eA-E])"


# extract answers with answer number and questions and options
def process_docx(path, output_path):
    doc = Document(path)
    full_text = "\n".join(p.text for p in doc.paragraphs)

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
