from flask import Flask, request, send_file, jsonify
import requests, os, re
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
from io import BytesIO

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        pptx_url = request.json.get("pptx_url")
        qr_url = request.json.get("qr_url")

        if not pptx_url or not qr_url:
            return jsonify({"error": "Missing pptx_url or qr_url"}), 400

        output_dir = "/tmp"
        qr_filename = os.path.basename(qr_url.split('?')[0])
        clean_name = re.sub(r'(?i)^qr[-_ ]*', '', qr_filename).replace('.png', '')

        pptx_path = os.path.join(output_dir, "template.pptx")
        qr_path = os.path.join(output_dir, "qr.png")
        output_pptx_path = os.path.join(output_dir, f"Kit de Bienvenida - {clean_name}.pptx")

        headers = {'User-Agent': 'Mozilla/5.0'}

        # Download PPTX
        r = requests.get(pptx_url, headers=headers)
        with open(pptx_path, 'wb') as f:
            f.write(r.content)

        # Download QR and validate
        r = requests.get(qr_url, headers=headers)
        image = Image.open(BytesIO(r.content))
        image.save(qr_path)

        # Inject QR
        prs = Presentation(pptx_path)
        inserted = False
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame and "{{QR_CODE}}" in shape.text:
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    slide.shapes._spTree.remove(shape._element)
                    slide.shapes.add_picture(qr_path, left, top, width=width, height=height)
                    inserted = True

        if not inserted:
            return jsonify({"error": "QR_CODE placeholder not found"}), 400

        prs.save(output_pptx_path)

        return send_file(output_pptx_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET"])
def home():
    return "✅ Crediviva PPTX QR API is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
