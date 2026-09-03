from django.http import HttpResponse
from django.shortcuts import render
from django.conf import settings
from django.contrib.auth.decorators import login_required
import os

from .models import GeneratedPaper
from .utils import generate_question_paper, generate_pdf, generate_answer_pdf, generate_docx


QUESTIONS_FLOW = [
    {'key': 'college_name', 'question': 'What is your college name? (Full name as it should appear on the paper)'},
    {'key': 'course_name', 'question': 'What is the course/degree name? (e.g. BACHELOR OF COMPUTER APPLICATIONS)'},
    {'key': 'subject', 'question': 'What is the subject name?'},
    {'key': 'subject_code', 'question': 'What is the subject code? (e.g. 23UCA506)'},
    {'key': 'semester', 'question': 'Which semester? (e.g. SEM-V)'},
    {'key': 'exam_type', 'question': 'What type of exam? (e.g. Internal Assessment Examination (IAE) - Set-I)'},
    {'key': 'exam_date', 'question': 'What is the exam date? (e.g. 29-07-2026)'},
    {'key': 'topics', 'question': 'Which topics/units should be covered? (comma separated)'},
    {'key': 'difficulty', 'question': 'What difficulty level? (type: Easy, Medium, or Hard)'},
]


@login_required(login_url='/accounts/login/')
def generate_paper_form(request):
    return render(request, 'chat_start.html')


@login_required(login_url='/accounts/login/')
def chat_reset(request):
    request.session['chat_answers'] = {}
    request.session['chat_step'] = 0
    return chat_view(request)


@login_required(login_url='/accounts/login/')
def chat_view(request):
    if request.method == 'POST':
        step = request.session.get('chat_step', 0)
        answers = request.session.get('chat_answers', {})

        if step < 0 or step >= len(QUESTIONS_FLOW):
            request.session['chat_answers'] = {}
            request.session['chat_step'] = 0
            return chat_view(request)

        current_key = QUESTIONS_FLOW[step]['key']
        answers[current_key] = request.POST.get('answer', '')

        request.session['chat_answers'] = answers
        step += 1
        request.session['chat_step'] = step
    else:
        step = request.session.get('chat_step', 0)
        answers = request.session.get('chat_answers', {})

    if step >= len(QUESTIONS_FLOW):
        return render(request, 'chat_summary.html', {'answers': answers})

    current_question = QUESTIONS_FLOW[step]['question']

    return render(request, 'chat_question.html', {
        'current_question': current_question,
        'step_number': step + 1,
        'total_steps': len(QUESTIONS_FLOW),
        'dots_range': range(1, len(QUESTIONS_FLOW) + 1),
    })


@login_required(login_url='/accounts/login/')
def chat_finish(request):
    answers = request.session.get('chat_answers', {})

    college_name = answers.get('college_name', '')
    course_name = answers.get('course_name', '')
    subject = answers.get('subject', '')
    subject_code = answers.get('subject_code', '')
    semester = answers.get('semester', '')
    exam_type = answers.get('exam_type', '')
    exam_date = answers.get('exam_date', '')
    topics = answers.get('topics', '')
    difficulty = answers.get('difficulty', '')

    data = generate_question_paper(subject, topics, difficulty)
    
    import json

    print("\n========== AI RESPONSE ==========")
    print(json.dumps(data, indent=4, ensure_ascii=False))
    print("=================================\n")
    header_info = {
        'college_name': college_name,
        'course_name': course_name,
        'subject': subject,
        'subject_code': subject_code,
        'semester': semester,
        'exam_type': exam_type,
        'exam_date': exam_date,
    }

    media_folder = os.path.join(settings.BASE_DIR, 'media', 'generated_papers')
    os.makedirs(media_folder, exist_ok=True)

    pdf_path = os.path.join(media_folder, f'question_paper_{subject}.pdf')
    answer_pdf_path = os.path.join(media_folder, f'answer_key_{subject}.pdf')
    docx_path = os.path.join(media_folder, f'question_paper_{subject}.docx')

    generate_pdf(data, header_info, pdf_path)
    generate_answer_pdf(data, header_info, answer_pdf_path)
    generate_docx(data, header_info, docx_path)

    paper = GeneratedPaper.objects.create(
        teacher=request.user,
        subject_name=subject,
        topics=topics,
        exam_type=exam_type,
        total_marks=50,
        no_of_questions=17,
        difficulty=difficulty,
        generated_content=str(data),
    )

    request.session['chat_answers'] = {}
    request.session['chat_step'] = 0

    return render(request, 'chat_result.html', {
        'subject': subject,
        'paper_id': paper.id,
        'pdf_filename': f'question_paper_{subject}.pdf',
        'answer_pdf_filename': f'answer_key_{subject}.pdf',
        'docx_filename': f'question_paper_{subject}.docx',
    })