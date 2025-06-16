from flask import Flask, render_template, request, make_response
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import cohere

app = Flask(__name__)
co = cohere.Client("NqCDyPmfZHiXEDiyn0Xooutz67b0XHFPoeZ8qeYy")  # Replace with your actual key

generated_data = {}

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    global generated_data
    data = request.form.to_dict()

    prompt = (
        f"Create a clean, professional resume for {data.get('name')} applying as a {data.get('job')}.\n"
        f"Email: {data.get('email')}\n"
        f"Phone: {data.get('phone')}\n"
        f"Summary: {data.get('summary')}\n"
        f"Skills: {data.get('skills')}\n"
        f"Work Experience: {data.get('experience')}\n"
        f"Education: {data.get('education')}\n"
        f"Certifications: {data.get('certifications', '')}\n"
        "Include the following sections: Summary, Skills, Work Experience, Education, Certifications.\n"
        "Do not include references or generic advice messages. Format with professional headings."
    )

    response = co.generate(model="command", prompt=prompt, max_tokens=1500, temperature=0.7)
    resume_text = response.generations[0].text.strip()

    # Remove unwanted phrases
    cleanup_phrases = [
        "Here is a sample resume",
        "REFERENCES\nAvailable upon request.",
        "Do you need more information added?",
        "Let me know what you would like to add or adjust.",
        "feel free to provide any additional information you would like to incorporate into your resume."
    ]
    for phrase in cleanup_phrases:
        resume_text = resume_text.replace(phrase, "")
    resume_text = resume_text.strip()

    generated_data = {"resume": resume_text, **data}
    return render_template("result.html", resume=resume_text, hide_buttons=False, **data)

@app.route("/download-word")
def download_word():
    global generated_data
    doc = Document()

    def add_heading(text, level=1, center=False):
        h = doc.add_heading(text, level)
        if center:
            h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    def add_paragraph(text, bold=False):
        if text.upper() == "N/A":
            return  # Don't add N/A text
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.font.size = Pt(11)

    # Header
    add_heading(generated_data.get("name", "Your Name"), level=0, center=True)
    add_paragraph(f"Job Title: {generated_data.get('job', '')}", bold=True)
    doc.add_paragraph()
    add_paragraph(f"Email: {generated_data.get('email', '')}")
    add_paragraph(f"Phone: {generated_data.get('phone', '')}")
    doc.add_paragraph()

    # Resume Content
    for line in generated_data.get("resume", "").split("\n"):
        line = line.strip()
        if not line:
            continue

        is_heading = line.endswith(":") or (line.isupper() and len(line.split()) <= 3)
        if is_heading:
            add_heading(line, level=2)
        elif line.upper() != "N/A":
            add_paragraph(line)

    # Word download
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    response = make_response(buffer.read())
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response.headers["Content-Disposition"] = "attachment; filename=resume.docx"
    return response

if __name__ == "__main__":
    app.run(debug=True)

