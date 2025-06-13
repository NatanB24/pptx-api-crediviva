import io
import requests
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from pptx import Presentation
from PIL import Image
import fitz  # PyMuPDF
import os

app = FastAPI()

@app.post("/pptx-api-crediviva")
async def insert_qr(request: Request):
    data = await request.json()
    pptx_url = data["pptx_url"]
    qr_url = data["qr_url"]

    # Download files
    pptx_response = requests.get(pptx_url)
    qr_response = requests.get(qr_url)
    pptx_bytes = io.BytesIO(pptx_response.content)
    qr_img = Image.open(io.BytesIO(qr_response.content))

    # Save QR image to disk
    qr_path = "qr.png"
    qr_img.save(qr_path)

    # Open PowerPoint
    prs = Presentation(pptx_bytes)
    slide = prs.slides[0]
    slide.shapes.add_picture(qr_path, left=0, top=0, width=prs.slide_width // 4, height=prs.slide_height // 4)

    # Save PowerPoint
    output_pptx_path = "output.pptx"
    prs.save(output_pptx_path)

    # Convert to PDF using LibreOffice
    os.system(f'libreoffice --headless --convert-to pdf "{output_pptx_path}" --outdir .')
    output_pdf_path = "output.pdf"

    # Prepare multipart response
    def file_iter():
        yield b"--boundary\n"
        yield b'Content-Disposition: form-data; name="pptx"; filename="output.pptx"\n'
        yield b"Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\n\n"
        yield open(output_pptx_path, "rb").read()
        yield b"\n--boundary\n"
        yield b'Content-Disposition: form-data; name="pdf"; filename="output.pdf"\n'
        yield b"Content-Type: application/pdf\n\n"
        yield open(output_pdf_path, "rb").read()
        yield b"\n--boundary--"

    return StreamingResponse(
        content=file_iter(),
        media_type="multipart/form-data; boundary=boundary"
    )
