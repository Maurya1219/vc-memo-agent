import io

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
from generate_docx import generate_memo_docx
from rag import ingest_pdf, ingest_excel, ask_question, generate_memo

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

class Question(BaseModel):
    question: str

@app.get("/", response_class=HTMLResponse)
async def root():
    with open("static/index.html") as f:
        return f.read()

@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    name = file.filename
    if not (name.endswith(".pdf") or name.endswith(".xlsx") or name.endswith(".xls")):
        raise HTTPException(status_code=400, detail="Only PDF and Excel files supported")
    contents = await file.read()
    tmp_path = f"/tmp/{name}"
    with open(tmp_path, "wb") as f:
        f.write(contents)
    count = ingest_pdf(tmp_path) if name.endswith(".pdf") else ingest_excel(tmp_path)
    return {"message": f"Ingested {count} chunks from {name}"}

@app.post("/ask")
async def ask(body: Question):
    if not body.question.strip():
        raise HTTPException(status_code=400, detail="Question cannot be empty")
    return ask_question(body.question)

@app.post("/generate-memo")
async def generate():
    return generate_memo()

@app.post("/generate-memo-docx")
async def generate_docx_endpoint():
    result = generate_memo()
    if "answer" not in result:
        raise HTTPException(status_code=500, detail="Memo generation failed")

    company_name = "Portfolio Company"
    docx_bytes = generate_memo_docx(result["answer"], company_name)

    return StreamingResponse(
        io.BytesIO(docx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=investment_memo.docx"}
    )
