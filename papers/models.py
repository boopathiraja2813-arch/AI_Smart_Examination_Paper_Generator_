from django.db import models
from django.contrib.auth.models import User

class GeneratedPaper(models.Model):
    teacher = models.ForeignKey(User, on_delete=models.CASCADE, null=True, blank=True)
    subject_name = models.CharField(max_length=200)
    topics = models.TextField()
    exam_type = models.CharField(max_length=50)
    total_marks = models.IntegerField()
    no_of_questions = models.IntegerField()
    difficulty = models.CharField(max_length=20)
    generated_content = models.TextField()
    pdf_file = models.FileField(upload_to='generated_papers/pdf/', blank=True, null=True)
    docx_file = models.FileField(upload_to='generated_papers/docx/', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.subject_name} - {self.exam_type}"