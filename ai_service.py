
import os
import json
from dotenv import load_dotenv
from google import genai

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("Gemini API key not found in .env")

client = genai.Client(api_key=API_KEY)

MODEL = "gemini-3.5-flash-lite"


def generate_text(prompt):

    response = client.models.generate_content(
        model=MODEL,
        contents=prompt
    )

    return response.text


def generate_questions(material_text):

    prompt = f"""
    You are an educational MCQ generation assistant.

    Read all the supplied study material carefully.

    Generate exactly 30 multiple-choice questions.

    STRICT SOURCE RESTRICTIONS:
    - Use ONLY the supplied study material.
    - Do not use external knowledge or information from memory.
    - Every question must be answerable using the supplied material.
    - Every correct answer must be supported by the material.
    - Explanations must also be based on the material.
    - Do not invent facts, examples or definitions.
    - Cover all important topics.
    - Avoid duplicate questions.

    MCQ REQUIREMENTS:
    - Each question must have exactly 4 options.
    - Include the correct answer.
    - Include a short explanation.
    - Assign a relevant topic to each question.
    - Include a mix of easy, medium and difficult questions.

    RETURN VALID JSON IN THIS EXACT STRUCTURE:

    {{
        "mcqs": [
            {{
                "topic": "Topic name",
                "question": "Question",
                "options": ["A", "B", "C", "D"],
                "answer": "Correct option",
                "explanation": "Explanation based on the material"
            }}
        ]
    }}

    Return JSON only. Do not use Markdown code fences.

    STUDY MATERIAL:
    {material_text[:25000]}
    """

    response = client.models.generate_content(
        model=MODEL,
        contents=prompt
    )

    questions = json.loads(response.text)

    if len(questions["mcqs"]) != 30:
        raise ValueError("Gemini did not generate exactly 30 MCQs.")

    for question in questions["mcqs"]:

        if len(question["options"]) != 4:
            raise ValueError("Each MCQ must have exactly 4 options.")

        if not question["topic"] or not question["answer"]:
            raise ValueError("MCQ topic or answer is missing.")

    return questions
