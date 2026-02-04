import json
import os
import sys
import re

# PYTHONPATH 추가 (패키지 인식 보장)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

from mcp.server.fastmcp import FastMCP

def log(msg):
    print(f"[*] {msg}", file=sys.stderr)

try:
    from cantops_email.client import EmailClient
    from cantops_email import document_loader
except ImportError:
    try:
        from client import EmailClient
        import document_loader
    except ImportError as e:
        log(f"CRITICAL IMPORT ERROR: {e}")

mcp = FastMCP("cts-email")

def _get_client():
    user = os.environ.get("EMAIL_USER")
    pw = os.environ.get("EMAIL_PASSWORD")
    if not user or not pw:
        return None, "❌ 설정 오류: EMAIL_USER, EMAIL_PASSWORD를 확인해주세요."
    try:
        return EmailClient(), None
    except Exception as e:
        return None, f"❌ 초기화 실패: {str(e)}"

def _format_email_list_table(emails):
    if not emails: return "📭 메일이 없습니다."
    table = "| ID | 날짜 | 보낸 사람 | 제목 |\n|---:|:---|:---|:---|
"
    for e in emails:
        d = str(e.get('date', ''))[:16]
        f = str(e.get('from', '')).replace('|', '&#124;')
        s = str(e.get('subject', '')).replace('|', '&#124;')
        table += f"| {e['id']} | {d}.. | {f} | {s} |\n"
    return table

@mcp.tool()
def list_emails(count: int = 10):
    """최근 이메일 목록을 조회합니다."""
    client, error = _get_client()
    if error: return error
    try:
        return _format_email_list_table(client.list_emails(count))
    except Exception as e: return f"❌ 에러: {str(e)}"

@mcp.tool()
def search_emails(keyword: str, limit: int = 50):
    """키워드로 메일을 검색합니다."""
    client, error = _get_client()
    if error: return error
    try:
        res = client.search_emails(keyword, limit)
        if not res: return f"🔍 '{keyword}' 검색 결과 없음."
        return f"🔍 **'{keyword}' 검색 결과**\n\n" + _format_email_list_table(res)
    except Exception as e: return f"❌ 에러: {str(e)}"

@mcp.tool()
def read_email(email_id: int):
    """메일 본문을 읽습니다. (ID 지정 필수)"""
    client, error = _get_client()
    if error: return error
    try:
        data = client.get_email(email_id)
        if not data: return "❌ 메일을 찾을 수 없음."
        att = ", ".join([f'`{f}`' for f in data.get('attachments', [])]) or "(없음)"
        return f"\n# 📧 ID: {data['id']}\n- From: {data['from']}\n- Date: {data['date']}\n- Attach: {att}\n\n## 본문\n---\n{data['body']}\n---\n"
    except Exception as e: return f"❌ 에러: {str(e)}"

@mcp.tool()
def send_email(to_email: str, subject: str, body: str):
    """이메일을 발송합니다. (사용자 승인 후 실행)"""
    client, error = _get_client()
    if error: return error
    try:
        return client.send_email(to_email, subject, body)
    except Exception as e: return f"❌ 에러: {str(e)}"

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """첨부파일을 다운로드합니다."""
    client, error = _get_client()
    if error: return error
    try:
        return client.download_file(email_id, filename, save_path)
    except Exception as e: return f"❌ 에러: {str(e)}"

@mcp.tool()
def read_document(file_path: str):
    """로컬 문서 파일을 읽습니다."""
    if not os.path.exists(file_path): return "❌ 파일 없음."
    try:
        content = document_loader.extract_text_from_file(file_path)
        return f"## 📄 {os.path.basename(file_path)}\n\n{content}"
    except Exception as e: return f"❌ 에러: {str(e)}"

def main():
    mcp.run()

if __name__ == "__main__":
    main()
