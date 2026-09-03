from google import genai
from decouple import config
import json

client = genai.Client(api_key=config('GEMINI_API_KEY'))

prompt = """
You are an expert exam question paper setter.

Generate a question paper with these details:
- Subject: Data Structures
- Topics: Linked Lists
- Exam type: CAT-1
- Total marks: 50
- Number of questions: 5
- Marks per question: 10
- Difficulty: Medium

Respond ONLY in valid JSON format, with NO extra text, NO markdown, NO explanation.
Use this exact structure:

{
  "questions": [
    {"q_no": 1, "question": "question text here", "marks": 10},
    {"q_no": 2, "question": "question text here", "marks": 10}
  ],
  "answers": [
    {"q_no": 1, "answer": "answer text here"},
    {"q_no": 2, "answer": "answer text here"}
  ]
}
"""

response = client.models.generate_content(
    model="gemini-3.5-flash",
    contents=prompt
)

raw_text = response.text.strip()
raw_text = raw_text.replace("```json", "").replace("```", "").strip()

data = json.loads(raw_text)

print("QUESTIONS:")
for q in data["questions"]:
    print(f"Q{q['q_no']}. {q['question']} [{q['marks']} marks]")

print("\nANSWERS:")
for a in data["answers"]:
    print(f"A{a['q_no']}. {a['answer']}")