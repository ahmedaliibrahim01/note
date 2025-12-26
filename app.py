from flask import Flask, render_template, request, send_file
import os
from processor import process_excels

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    # Vardiya dosyasını al
    shift_file = request.files["shift_file"]
    shift_path = os.path.join(UPLOAD_FOLDER, shift_file.filename)
    shift_file.save(shift_path)
    
    # Note dosyasını al (opsiyonel)
    note_path = None
    if "note_file" in request.files and request.files["note_file"].filename != "":
        note_file = request.files["note_file"]
        note_path = os.path.join(UPLOAD_FOLDER, note_file.filename)
        note_file.save(note_path)
    
    # Tarih bilgilerini al
    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")

    # Process et
    output_path = process_excels(
        shift_path=shift_path,
        note_path=note_path,
        start_date=start_date,
        end_date=end_date
    )

    # Excel'i indir
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)