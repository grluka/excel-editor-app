from flask import Flask, render_template, request, send_file
import openpyxl
from PIL import Image
from openpyxl.drawing.image import Image as ExcelImage
import json
import os
from pdf2image import convert_from_path

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

with open("text_fields.json", "r") as f:
    text_fields = json.load(f)

with open("image_fields.json", "r") as f:
    image_fields = json.load(f)

with open("multiple_images.json", "r") as f:
    multiple_images = json.load(f)

with open("combined_upload.json", "r") as f:
    combined_uploads = json.load(f)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        excel_file = request.files["excel_file"]
        if excel_file and excel_file.filename:
            file_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
            excel_file.save(file_path)
        else:
            return "Excel datoteka nije učitana", 400

        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        for field in text_fields:
            value = request.form.get(field["Polje"])
            if value:
                ws[field["Polje"]] = value

        for group in multiple_images:
            files = request.files.getlist(group["Opis"])
            polja = group["Polja"]
            for i, img_file in enumerate(files):
                if i >= len(polja):
                    break
                if img_file and img_file.filename:
                    filename = f"{group['Opis'].replace(' ', '_')}_{i}_{img_file.filename}"
                    img_path = os.path.join(UPLOAD_FOLDER, filename)
                    img_file.save(img_path)
                    if img_path.lower().endswith((".png", ".jpg", ".jpeg")):
                        img = ExcelImage(img_path)
                        img.width, img.height = group["Dimenzije u pixelima"]
                        ws.add_image(img, polja[i])

        for image in image_fields:
            if image["Polje"] in request.files:
                img_file = request.files[image["Polje"]]
                if img_file and img_file.filename:
                    img_path = os.path.join(UPLOAD_FOLDER, img_file.filename)
                    img_file.save(img_path)
                    if img_path.lower().endswith((".png", ".jpg", ".jpeg")):
                        img = ExcelImage(img_path)
                        img.width, img.height = image["Dimenzije u pixelima"]
                        ws.add_image(img, image["Polje"])

        for combo in combined_uploads:
            uploads = request.files.getlist(combo["Opis"])
            all_images = []
            for file in uploads:
                if file and file.filename:
                    ext = os.path.splitext(file.filename)[1].lower()
                    file_path = os.path.join(UPLOAD_FOLDER, f"combo_{file.filename}")
                    file.save(file_path)

                    if ext == ".pdf":
                        try:
                            pdf_images = convert_from_path(file_path)
                            for idx, img in enumerate(pdf_images):
                                image_path = os.path.join(UPLOAD_FOLDER, f"{file.filename}_{idx}.png")
                                img.save(image_path, "PNG")
                                all_images.append(image_path)
                        except Exception as e:
                            print(f"Greška kod konverzije PDF-a: {file.filename} -> {e}")
                    elif ext in [".png", ".jpg", ".jpeg"]:
                        all_images.append(file_path)

            for i, img_path in enumerate(all_images):
                if i >= len(combo["Polja"]):
                    break
                if img_path.lower().endswith((".png", ".jpg", ".jpeg")):
                    img = ExcelImage(img_path)
                    img.width, img.height = combo["Dimenzije u pixelima"]
                    ws.add_image(img, combo["Polja"][i])

        output_path = os.path.join(UPLOAD_FOLDER, "output.xlsx")
        wb.save(output_path)

        response = send_file(output_path, as_attachment=True)

        for f in os.listdir(UPLOAD_FOLDER):
            full_path = os.path.join(UPLOAD_FOLDER, f)
            if os.path.isfile(full_path):
                try:
                    os.remove(full_path)
                except Exception as e:
                    print(f"Ne mogu obrisati {f}: {e}")

        return response

    return render_template("template.html",
                           text_fields=text_fields,
                           image_fields=image_fields,
                           multiple_images=multiple_images,
                           combined_uploads=combined_uploads)

if __name__ == "__main__":
    app.run(debug=True)
