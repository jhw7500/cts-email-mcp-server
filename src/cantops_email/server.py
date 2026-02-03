import json
import os
from mcp.server.fastmcp import FastMCP
from .client import EmailClient
from . import document_loader

mcp = FastMCP("Cantops-Email")
client = EmailClient()

def _format_email_list_table(emails):
    """이메일 리스트를 Markdown 표로 변환"""
    if not emails:
        return "📭 조회된 이메일이 없습니다."
    
    table = "| ID | 날짜 | 보낸 사람 | 제목 |
"
    table += "|---:|:---|:---|:---|
"
    for e in emails:
        # 날짜와 내용의 특수문자 처리
        date_str = str(e.get('date', ''))[:16]  
        sender = str(e.get('from', '')).replace('|', '&#124;') 
        subject = str(e.get('subject', '')).replace('|', '&#124;')
        table += f"| {e['id']} | {date_str}.. | {sender} | {subject} |
"
    return table

@mcp.tool()
def list_emails(count: int = 10):
    """최근 이메일 목록을 깔끔한 표 형식으로 조회합니다."""
    emails = client.list_emails(count)
    return _format_email_list_table(emails)

@mcp.tool()
def read_email(email_id: int):
    """
    특정 이메일의 내용을 읽기 좋은 보고서 형식으로 보여줍니다.
    """
    email_data = client.get_email(email_id)
    if not email_data:
        return f"❌ ID {email_id}번 이메일을 찾을 수 없습니다."

    # 첨부파일 리스트 포맷팅
    attachments = email_data.get('attachments', [])
    att_str = ", ".join([f"`{f}`" for f in attachments]) if attachments else "(없음)"

    markdown_report = f"""
# 📧 이메일 상세 내용 (ID: {email_data['id']})

| 항목 | 내용 |
|---|---|
| **보낸 사람** | {email_data['from']} |
| **날짜** | {email_data['date']} |
| **제목** | {email_data['subject']} |
| **첨부파일** | {att_str} |

## 📝 본문 내용
---
{email_data['body']}
---
"""
    return markdown_report

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """첨부파일 다운로드 결과를 명확하게 알립니다."""
    result = client.download_file(email_id, filename, save_path)
    if result.startswith("Success"):
        # 성공 메시지에서 경로 추출
        saved_file = result.replace("Success: ", "")
        return f"✅ **다운로드 완료!**\n- 결과: `{saved_file}`"
    else:
        return f"❌ **다운로드 실패**\n- 원인: {result}"

@mcp.tool()
def read_document(file_path: str):
    """문서 내용을 읽어 Markdown 인용구로 보여줍니다."""
    if not os.path.exists(file_path):
        return f"❌ 파일이 존재하지 않습니다: `{file_path}`"
    
    content = document_loader.extract_text_from_file(file_path)
    if content.startswith("Error"):
        return f"❌ **문서 읽기 실패**: {content}"
    
    return f"## 📄 문서 내용: {os.path.basename(file_path)}\n\n```text\n{content}\n```"

def main(): mcp.run()