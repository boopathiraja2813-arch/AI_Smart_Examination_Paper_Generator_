from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# Sample data - AI kudutha structure mari (test ku)
sample_data = {
    "part_a": [
        {"q_no": i, "question": f"Sample MCQ question number {i}?",
         "options": {"A": "Option one", "B": "Option two", "C": "Option three", "D": "Option four"},
         "answer": "A"} for i in range(1, 11)
    ],
    "part_b": [
        {"q_no": i, "question": f"Sample short answer question number {i}?", "answer": "Sample answer text"} for i in range(1, 5)
    ],
    "part_c": [
        {"q_no": i, "question_a": f"Sample essay question {i}a?", "question_b": f"Sample essay question {i}b?",
         "answer_a": "Answer A", "answer_b": "Answer B"} for i in range(1, 4)
    ]
}

header_info = {
    'college_name': 'Excel College for Commerce and Science',
    'course_name': 'BACHELOR OF COMPUTER APPLICATIONS',
    'subject': 'Python Programming',
    'subject_code': '23UCA506',
    'semester': 'SEM-V',
    'exam_type': 'Internal Assessment Examination (IAE) - Set-I',
    'exam_date': '29-07-2026',
}


def split_text(text, max_chars):
    words = text.split()
    lines = []
    current = ""
    for word in words:
        if len(current) + len(word) + 1 <= max_chars:
            current += (" " if current else "") + word
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def format_options(options):
    parts = []
    for key, value in options.items():
        value = (value or "").strip()
        if value:
            parts.append(f"({key}) {value}")
    return "   ".join(parts)


def draw_header(c, width, height, info):
    y = height - 50
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(width / 2, y, info['college_name'])
    y -= 22

    c.setFont("Helvetica", 10)
    c.drawCentredString(width / 2, y, info['course_name'])
    y -= 16

    c.setFont("Helvetica", 11)
    c.drawCentredString(width / 2, y, info['exam_type'])
    y -= 18

    c.setFont("Helvetica", 10)
    c.drawCentredString(width / 2, y, f"{info['subject_code']}  -  {info['subject']}")
    y -= 16

    c.drawCentredString(width / 2, y, f"{info['semester']}      Date: {info['exam_date']}      Maximum Marks: 50")
    y -= 20

    c.line(40, y, width - 40, y)
    y -= 20
    return y


def check_page_break(c, y, width, height, info, min_space=60):
    if y < min_space:
        c.showPage()
        y = draw_header(c, width, height, info)
    return y


def draw_wrapped(c, x, y, text, size=10, max_chars=95, line_gap=14):
    c.setFont("Helvetica", size)
    lines = split_text(text, max_chars)
    for line in lines:
        c.drawString(x, y, line)
        y -= line_gap
    return y


def generate_test_pdf(data, info, filepath):
    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4

    y = draw_header(c, width, height, info)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "PART - A")
    c.drawRightString(width - 40, y, "(10 x 1 = 10 Marks)")
    y -= 10
    c.line(40, y, width - 40, y)
    y -= 18

    for q in data["part_a"]:
        y = check_page_break(c, y, width, height, info)
        y = draw_wrapped(c, 40, y, f"{q['q_no']}. {q['question']}")
        opt_line = format_options(q['options'])
        y = check_page_break(c, y, width, height, info)
        y = draw_wrapped(c, 55, y, opt_line, max_chars=100)
        y -= 8

    y -= 10
    y = check_page_break(c, y, width, height, info)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "PART - B")
    c.drawRightString(width - 40, y, "(Answer any TWO) (2 x 5 = 10 Marks)")
    y -= 10
    c.line(40, y, width - 40, y)
    y -= 18

    for q in data["part_b"]:
        y = check_page_break(c, y, width, height, info)
        y = draw_wrapped(c, 40, y, f"{q['q_no']}. {q['question']}")
        y -= 10

    y -= 10
    y = check_page_break(c, y, width, height, info)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "PART - C")
    c.drawRightString(width - 40, y, "(3 x 10 = 30 Marks)")
    y -= 10
    c.line(40, y, width - 40, y)
    y -= 18

    for idx, q in enumerate(data["part_c"], start=1):
        y = check_page_break(c, y, width, height, info)
        y = draw_wrapped(c, 40, y, f"{idx}. a) {q['question_a']}")
        y -= 4
        c.setFont("Helvetica", 10)
        c.drawCentredString(width / 2, y, "(OR)")
        y -= 16
        y = check_page_break(c, y, width, height, info)
        y = draw_wrapped(c, 40, y, f"    b) {q['question_b']}")
        y -= 14

    c.save()
    print(f"PDF created: {filepath}")


generate_test_pdf(sample_data, header_info, "test_partabc_output.pdf")