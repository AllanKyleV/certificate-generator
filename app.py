from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import io
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():

    names = request.form.getlist('name')
    directors = request.form.getlist('director')
    coaches = request.form.getlist('coach')

    people = []

    for i in range(len(names)):
        people.append({
            "name": names[i],
            "director": directors[i],
            "coach": coaches[i]
        })

    context = {
        "people": people
    }

    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "template.docx")

    doc = DocxTemplate(template_path)
    doc.render(context)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="certificates.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    app.run(debug=True)