import uuid
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import fitz  #  PyMuPDF
import os
import io
import zipfile
from docx import Document
from PIL import Image
import asyncio

app = FastAPI()
@app.get("/")
def root():
	return {"message": "Welcome to the XPS API. Use /docs for API documentation."}
# Configuration
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# CORS Setup
app.add_middleware(
	CORSMiddleware,
	allow_origins=["*"],
	allow_methods=["*"],
	allow_headers=["*"],
)

# Supported conversions
SUPPORTED_CONVERSIONS = {
	"pdf": "application/pdf",
	"images": "application/zip",
	"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
}

# Helper Functions
async def save_file(file: UploadFile) -> str:
	"""Save uploaded file with a UUID filename."""
	safe_filename = f"{uuid.uuid4()}.xps"
	input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
	with open(input_path, "wb") as f:
		f.write(await file.read())
	return input_path

def cleanup(file_path: str):
	"""Remove temporary files."""
	if os.path.exists(file_path):
		os.remove(file_path)

# API Endpoints
@app.post("/convert/{conversion_type}")
async def convert_xps(conversion_type: str, file: UploadFile = File(...)):
	"""Convert XPS to PDF, images (ZIP), or DOCX."""
	if conversion_type not in SUPPORTED_CONVERSIONS:
		raise HTTPException(status_code=400, detail=f"Unsupported conversion. Allowed: {list(SUPPORTED_CONVERSIONS.keys())}")

	if not file.filename.lower().endswith(".xps"):
		raise HTTPException(status_code=400, detail="Only .xps files are allowed")

	input_path = await save_file(file)
	doc = None
	try:
		doc = fitz.open(input_path)
		output_filename = f"{uuid.uuid4()}.{conversion_type}"
		output_path = os.path.join(RESULT_FOLDER, output_filename)

		if conversion_type == "pdf":
			await asyncio.to_thread(doc.save, output_path)
			return FileResponse(output_path, media_type=SUPPORTED_CONVERSIONS["pdf"], filename=output_filename)

		elif conversion_type == "images":
			zip_buffer = io.BytesIO()
			with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zipf:
				for i in range(len(doc)):
					page = doc.load_page(i)
					pix = page.get_pixmap()
					img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
					img_io = io.BytesIO()
					img.save(img_io, format="PNG")
					img_io.seek(0)
					zipf.writestr(f"page_{i+1}.png", img_io.read())
			zip_buffer.seek(0)
			return StreamingResponse(
				zip_buffer,
				media_type="application/zip",
				headers={"Content-Disposition": "attachment; filename=images.zip"}
			)

		elif conversion_type == "docx":
			word_doc = Document()
			for i in range(len(doc)):
				text = doc.load_page(i).get_text()
				word_doc.add_paragraph(text)
				if i < len(doc) - 1:
					word_doc.add_page_break()
			word_doc.save(output_path)
			return FileResponse(output_path, media_type=SUPPORTED_CONVERSIONS["docx"], filename=output_filename)

	except Exception as e:
		raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
	finally:
		if doc:
			doc.close()
		cleanup(input_path)

@app.post("/read-xps")
async def read_xps(file: UploadFile = File(...)):
	"""Extract text and metadata from XPS."""
	if not file.filename.lower().endswith(".xps"):
		raise HTTPException(status_code=400, detail="Only .xps files are allowed")

	input_path = await save_file(file)
	doc = None
	try:
		doc = fitz.open(input_path)
		metadata = doc.metadata
		text_content = ""
        
		for page_num in range(len(doc)):
			page = doc.load_page(page_num)
			text_content += page.get_text() + "\n\n--- Page Break ---\n\n"

		return JSONResponse({
			"metadata": metadata,
			"page_count": len(doc),
			"text": text_content
		})
	except Exception as e:
		raise HTTPException(status_code=500, detail=f"Failed to read XPS: {str(e)}")
	finally:
		if doc:
			doc.close()
		cleanup(input_path)

@app.post("/preview-all")
async def preview_all_xps(file: UploadFile = File(...)):
	"""Generate PNG previews of all XPS pages (ZIP)."""
	if not file.filename.lower().endswith(".xps"):
		raise HTTPException(status_code=400, detail="Only .xps files are allowed")

	input_path = await save_file(file)
	doc = None
	try:
		doc = fitz.open(input_path)
		zip_buffer = io.BytesIO()
        
		with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zipf:
			for i in range(len(doc)):
				page = doc.load_page(i)
				pix = page.get_pixmap()
				img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
				img_io = io.BytesIO()
				img.save(img_io, format="PNG")
				img_io.seek(0)
				zipf.writestr(f"page_{i+1}.png", img_io.read())
        
		zip_buffer.seek(0)
		return StreamingResponse(
			zip_buffer,
			media_type="application/zip",
			headers={"Content-Disposition": "attachment; filename=preview_pages.zip"}
		)
	except Exception as e:
		raise HTTPException(status_code=500, detail=f"Failed to generate preview: {str(e)}")
	finally:
		if doc:
			doc.close()
		cleanup(input_path)

@app.post("/edit-xps")
async def edit_xps(
	file: UploadFile = File(...),
	old_text: str = Form(...),
	new_text: str = Form(...)
):
	"""Replace text in XPS and save as PDF."""
	if not file.filename.lower().endswith(".xps"):
		raise HTTPException(status_code=400, detail="Only .xps files are allowed")

	input_path = await save_file(file)
	doc = None
	try:
		doc = fitz.open(input_path)
        
		for page_num in range(len(doc)):
			page = doc.load_page(page_num)
			text_instances = page.search_for(old_text)
            
			for inst in text_instances:
				page.add_redact_annot(inst)
				page.apply_redactions()
				page.insert_text(
					(inst.x0, inst.y0),
					new_text,
					fontsize=11,
					color=(0, 0, 0)
				)

		output_filename = f"edited_{uuid.uuid4()}.pdf"
		output_path = os.path.join(RESULT_FOLDER, output_filename)
		doc.save(output_path)
		return FileResponse(
			output_path,
			media_type="application/pdf",
			filename=output_filename
		)
	except Exception as e:
		raise HTTPException(status_code=500, detail=f"Failed to edit XPS: {str(e)}")
	finally:
		if doc:
			doc.close()
		cleanup(input_path)
