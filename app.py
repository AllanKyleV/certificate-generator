from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import io
import os
import csv

# --------------------------------------------------
# Create Flask App (MUST be before any @app.route)
# --------------------------------------------------
app = Flask(__name__)


# --------------------------------------------------
# Home Page
# --------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')


# --------------------------------------------------
# Generate Certificates
# --------------------------------------------------
@app.route('/generate', methods=['POST'])
def generate():

    people = []

    # Check if CSV file was uploaded
    csv_file = request.files.get('csv_file')

    if csv_file and csv_file.filename != "":
        try:
            stream = csv_file.stream.read().decode("UTF8").splitlines()
            reader = csv.DictReader(stream)

            # Validate required columns
            required_fields = {"name", "director", "coach"}

            if not reader.fieldnames or not required_fields.issubset(reader.fieldnames):
                return "CSV must contain columns: name, director, coach"

            for row in reader:
                if row["name"].strip():
                    people.append({
                        "name": row["name"],
                        "director": row["director"],
                        "coach": row["coach"]
                    })

        except Exception as e:
            return f"Error reading CSV: {str(e)}"

    else:
        # Manual entry mode
        names = request.form.getlist('name')
        directors = request.form.getlist('director')
        coaches = request.form.getlist('coach')

        for i in range(len(names)):
            if names[i].strip():
                people.append({
                    "name": names[i],
                    "director": directors[i],
                    "coach": coaches[i]
                })

    if not people:
        return "No valid data provided."

    # --------------------------------------------------
    # Load Template
    # --------------------------------------------------
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "template.docx")

    if not os.path.exists(template_path):
        return "template.docx not found in project folder."

    doc = DocxTemplate(template_path)

    context = {
        "people": people
    }

    doc.render(context)

    # Save to memory
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="certificates.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# --------------------------------------------------
# Run App
# --------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True)