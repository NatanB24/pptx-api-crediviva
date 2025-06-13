from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.util import Inches
import requests
import tempfile
import os

app = FastAPI()

@app.post("/pptx-api-crediviva")
async def generate_pptx(request: Request):
    data = await request.json()
    qr_url = data.get("qr_url")
    template_url = data.get("template_url")

    if not qr_url or not template_url:
        return {"error": "Missing 'qr_url' or 'template_url'"}

    # Download template
    template_response = requests.get(template_url)
    if template_response.status_code != 200:
        return {"error": "Failed to download template"}
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_template:
        tmp_template.write(template_response.content)
        tmp_template_path = tmp_template.name

    # Download QR code
    qr_response = requests.get(qr_url)
    if qr_response.status_code != 200:
        return {"error": "Failed to download QR code"}
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_qr:
        tmp_qr.write(qr_response.content)
        tmp_qr_path = tmp_qr.name

    # Open presentation and insert QR
    prs = Presentation(tmp_template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and "{{ QR_CODE }}" in shape.text:
                left = shape.left
                top = shape.top
                height = shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(tmp_qr_path, left, top, height=height)

    # Save updated presentation
    output_path = os.path.join(tempfile.gettempdir(), "output.pptx")
    prs.save(output_path)

    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="Crediviva_Template_Output.pptx")
