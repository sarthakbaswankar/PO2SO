from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from app.converter import html_to_pdf
import io
import uuid
from datetime import datetime

app = FastAPI(
    title="HTML to PDF Converter",
    description="Converts Microsoft Graph API HTML email bodies to PDF.",
    version="1.0.0",
)


class HtmlInput(BaseModel):
    content: str

    class Config:
        json_schema_extra = {
            "example": {
                "content": "<html><body><p>Hello World</p></body></html>"
            }
        }


@app.get("/")
def root():
    return {"message": "HTML to PDF API is running. POST /convert with { \"content\": \"...\" }"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/convert")
def convert(payload: HtmlInput):
    if not payload.content or not payload.content.strip():
        raise HTTPException(status_code=400, detail="content field must not be empty.")

    # Generate unique filename: email_20260601_143022_a1b2c3d4.pdf
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    short_id = str(uuid.uuid4()).replace("-", "")[:8]
    file_id = f"{timestamp}_{short_id}"

    try:
        pdf_bytes = html_to_pdf(payload.content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={
            "Content-Disposition": f"attachment; filename=email_{file_id}.pdf",
            "X-File-ID": file_id        # also exposed in response headers if needed
        },
    )