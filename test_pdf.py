from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def generate_pdf(questions, filepath):
    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, height-50, "XYZ College of Engineering")

    c.setFont("Helvetica", 12)
    c.drawCentredString(width/2, height-70, "CAT-1 - Data Structures")

    y = height - 110
    c.setFont("Helvetica", 11)
    for q in questions:
        text = f"Q{q['q_no']}. {q['question']} [{q['marks']} marks]"
        lines = split_text(text, 90)
        for line in lines:
            c.drawString(50, y, line)
            y -= 20
        y -= 10

    c.save()

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


sample_questions = [
    {"q_no": 1, "question": "Explain Floyd's Cycle Detection Algorithm to detect a loop in a linked list.", "marks": 10},
    {"q_no": 2, "question": "Write an algorithm to reverse a doubly linked list in-place.", "marks": 10},
]

generate_pdf(sample_questions, "sample_question_paper.pdf")
print("PDF generated successfully!")