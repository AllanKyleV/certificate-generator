# Load template
# Create context dictionary
# Render
# Save as certificate_output.docx

from docxtpl import DocxTemplate

# Load template
# doc = DocxTemplate("certificate_template.docx")
doc = DocxTemplate("certificate_template.docx")

# Hardcoded data (one-off)
# context = {
#    "name": "Guido Van Rossum",
#    "director": "Brendan Eich",
#    "coach": "Dennis Ritchie"
#}

context = {
    "people": [
        {
            "name": "Kyle Adrian Mendoza",
            "director": "Maria Theresa Santos",
            "coach": "Jonathan Reyes"
        },
        {
            "name": "Anna Beatriz Cruz",
            "director": "Maria Theresa Santos", 
            "coach": "Jonathan Reyes"
        },
        {
            "name": "Mark Anthony Villanueva",
            "director": "Daniel Roberto Lim",
            "coach": "Sophia Delgado"
        },
        {
            "name": "Leah Camille Navarro",
            "director": "Daniel Roberto Lim",
            "coach": "Sophia Delgado"
        },
        {
            "name": "Joshua Miguel Fernandez",
            "director": "Catherine Ramos",
            "coach": "Michael Torres"
        }
    ]
}

# Render template   
doc.render(context)

# Save output
doc.save("certificate_output.docx")

print("Certificate generated successfully.")