from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from fpdf import FPDF
from pdf2image import convert_from_path
from docx2pdf import convert as docx_to_pdf
from PIL import Image
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/convert', methods=['POST'])
def convert_file():
    file = request.files['file']
    conversion_type = request.form['conversion']

    if not file:
        return "No file uploaded", 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    output_filename = ""
    output_path = ""

    try:
        # 1️⃣ PDF → Word
        if conversion_type == "pdf_to_word":
            output_filename = filename.replace('.pdf', '.docx')
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

        # 2️⃣ Word → PDF
        elif conversion_type == "word_to_pdf":
            output_filename = filename.replace('.docx', '.pdf')
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            docx_to_pdf(input_path, output_path)

        # 3️⃣ Image → PDF
        elif conversion_type == "image_to_pdf":
            output_filename = os.path.splitext(filename)[0] + ".pdf"
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            img = Image.open(input_path).convert("RGB")
            img.save(output_path)

        # 4️⃣ PDF → Image
        elif conversion_type == "pdf_to_image":
            images = convert_from_path(input_path)
            output_folder = os.path.join(RESULT_FOLDER, filename.split('.')[0])
            os.makedirs(output_folder, exist_ok=True)
            for i, img in enumerate(images):
                img_path = os.path.join(output_folder, f"page_{i+1}.jpg")
                img.save(img_path, "JPEG")
            return f"PDF converted to images in folder: {output_folder}"

        # 5️⃣ Text → PDF
        elif conversion_type == "text_to_pdf":
            output_filename = filename.replace('.txt', '.pdf')
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            with open(input_path, "r", encoding="utf-8") as f:
                for line in f:
                    # Clean line to remove unsupported Unicode characters
                    clean_line = (
                        line.replace("–", "-")
                            .replace("—", "-")
                            .replace("“", '"')
                            .replace("”", '"')
                            .replace("‘", "'")
                            .replace("’", "'")
                            .encode("latin-1", "ignore")
                            .decode("latin-1")
                    )
                    pdf.multi_cell(0, 10, clean_line)
            pdf.output(output_path)

        # 6️⃣ Excel → CSV
        elif conversion_type == "excel_to_csv":
            output_filename = filename.replace('.xlsx', '.csv')
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            df = pd.read_excel(input_path)
            df.to_csv(output_path, index=False)

        # 7️⃣ CSV → Excel
        elif conversion_type == "csv_to_excel":
            output_filename = filename.replace('.csv', '.xlsx')
            output_path = os.path.join(RESULT_FOLDER, output_filename)
            df = pd.read_csv(input_path)
            df.to_excel(output_path, index=False)

        else:
            return "Unsupported conversion", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error during conversion: {str(e)}", 500


if __name__ == "__main__":
    app.run(debug=True)
