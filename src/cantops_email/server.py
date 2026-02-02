import json
import os
from mcp.server.fastmcp import FastMCP
from .client import EmailClient
from . import document_loader

mcp = FastMCP("Cantops-Email")
client = EmailClient()

@mcp.tool()
def list_emails(count: int = 10):
    """최근 이메일 목록을 조회합니다."""
    return json.dumps(client.list_emails(count), ensure_ascii=False)

@mcp.tool()
def read_email(email_id: int):
    """특정 ID의 이메일 내용과 첨부파일 목록을 읽습니다."""
    return json.dumps(client.get_email(email_id), ensure_ascii=False)

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """이메일의 첨부파일을 지정된 경로에 다운로드합니다."""
    return client.download_file(email_id, filename, save_path)

@mcp.tool()
def read_document(file_path: str):
    """
    로컬 파일의 텍스트 내용을 읽어옵니다.
    지원 형식: .pptx, .docx, .xlsx, .pdf, .txt, .md, .py 등 소스코드
    """
    if not os.path.exists(file_path):
        return f"Error: 파일이 존재하지 않습니다: {file_path}"
    return document_loader.extract_text_from_file(file_path)

def main(): mcp.run()
