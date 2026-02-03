import json
import os
import sys
from mcp.server.fastmcp import FastMCP

# MCP 서버가 stdout을 사용하므로 로그는 반드시 stderr로 보내야 합니다.
def log(msg):
    print(f"[*] {msg}", file=sys.stderr)

try:
    from .client import EmailClient
    from . import document_loader
except ImportError as e:
    log(f"Import Error: {e}")
    # 상대 경로 임포트 실패 시 (로컬 실행 등) 대비
    try:
        from client import EmailClient
        import document_loader
    except ImportError:
        log("Critical: Could not load internal modules")

mcp = FastMCP("Cantops-Email")

def _get_client():
    """EmailClient 인스턴스를 안전하게 생성합니다."""
    try:
        user = os.environ.get("EMAIL_USER")
        pw = os.environ.get("EMAIL_PASSWORD")
        if not user or not pw:
            return None, "❌ 설정 오류: settings.json의 EMAIL_USER, EMAIL_PASSWORD를 확인해주세요."
        return EmailClient(), None
    except Exception as e:
        return None, f"❌ 클라이언트 초기화 에러: {str(e)}"

def _format_email_list_table(emails):
    if not emails:
        return "📭 조회된 이메일이 없습니다."
    
    table = "| ID | 날짜 | 보낸 사람 | 제목 |\n"
    table += "|---:|:---|:---|:---|" + "\n"
    for e in emails:
        date_str = str(e.get('date', ''))[:16]  
        sender = str(e.get('from', '')).replace('|', '&#124;') 
        subject = str(e.get('subject', '')).replace('|', '&#124;')
        table += f"| {e['id']} | {date_str}.. | {sender} | {subject} |\n"
    return table

@mcp.tool()
def list_emails(count: int = 10):
    """최근 이메일 목록을 깔끔한 표 형식으로 조회합니다."""
    client, error = _get_client()
    if error: return error
    
    try:
        emails = client.list_emails(count)
        return _format_email_list_table(emails)
    except Exception as e:
        return f"❌ 목록 조회 실패: {str(e)}"

@mcp.tool()
def read_email(email_id: int):
    """이메일 상세 내용을 보고서 형식으로 읽습니다."""
    client, error = _get_client()
    if error: return error

    try:
        email_data = client.get_email(email_id)
        if not email_data:
            return f"❌ ID {email_id}번 이메일을 찾을 수 없습니다."

        attachments = email_data.get('attachments', [])
        att_str = ", ".join([f'`{f}`' for f in attachments]) if attachments else "(없음)"

        return f"""
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
    except Exception as e:
        return f"❌ 메일 읽기 실패: {str(e)}"

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """첨부파일을 지정된 경로에 다운로드합니다."""
    client, error = _get_client()
    if error: return error

    try:
        result = client.download_file(email_id, filename, save_path)
        if result.startswith("Success"):
            saved_file = result.replace("Success: ", "")
            return f"✅ **다운로드 완료!**\n- 결과: `{saved_file}`"
        return f"❌ **다운로드 실패**: {result}"
    except Exception as e:
        return f"❌ 에러 발생: {str(e)}"

@mcp.tool()
def read_document(file_path: str):
    """문서(.pptx, .xlsx, .pdf 등) 내용을 텍스트로 읽습니다."""
    if not os.path.exists(file_path):
        return f"❌ 파일이 존재하지 않습니다: `{file_path}`"
    
    try:
        content = document_loader.extract_text_from_file(file_path)
        return f"## 📄 문서 내용: {os.path.basename(file_path)}\n\n```text\n{content}\n```"
    except Exception as e:
        return f"❌ 문서 읽기 실패: {str(e)}"

def main():
    mcp.run()

if __name__ == "__main__":
    main()