from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from typing import List
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

app = FastAPI()

@app.get("/")
def root():
    return {"status": "ok"}

@app.post("/generate_catalog")
async def generate_catalog(excel: UploadFile = File(...), images: List[UploadFile] = File(...)):
    wb = load_workbook(filename=io.BytesIO(await excel.read()))
    ws = wb.active

    for idx, image_file in enumerate(images, start=2):  # починаємо з другого рядка
        img_data = await image_file.read()
        image = Image.open(io.BytesIO(img_data))
        output = io.BytesIO()
        image.save(output, format='PNG')
        output.seek(0)
        img = XLImage(output)
        img.width = 100
        img.height = 100
        ws.add_image(img, f"B{idx}")

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return StreamingResponse(output_stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=catalog_ready.xlsx"})
