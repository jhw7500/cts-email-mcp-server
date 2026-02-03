import json
import os
import sys
import re
from mcp.server.fastmcp import FastMCP

# 모든 로그는 반드시 stderr로 출력하여 MCP 통신(stdout)을 방해하지 않아야 합니다.
def log(msg):
    print(f"[*] {msg}", file=sys.stderr)

# 모듈 임포트 시도 (패키지 설치 및 로컬 실행 대응)
try:
    from cantops_email.client import EmailClient
    from cantops_email import document_loader
except ImportError:
    try:
        from .client import EmailClient
        from . import document_loader
    except ImportError:
        try:
            from client import EmailClient
            import document_loader
        except ImportError:
            log("Critical: 내부 모듈을 불러올 수 없습니다.")

# MCP 서버 초기화
mcp = FastMCP("cts-email")

def _get_client():
    """EmailClient 인스턴스를 안전하게 생성합니다."""
    try:
        user = os.environ.get("EMAIL_USER")
        pw = os.environ.get("EMAIL_PASSWORD")
        if not user or not pw:
            return None, "❌ 설정 오류: settings.json에서 EMAIL_USER와 EMAIL_PASSWORD 환경 변수를 확인해주세요."
        return EmailClient(), None
    except Exception as e:
        return None, f"❌ 클라이언트 초기화 에러: {str(e)}"

def _format_email_list_table(emails):
    if not emails:
        return "📭 조회된 이메일이 없습니다."
    table = "| ID | 날짜 | 보낸 사람 | 제목 |\n|---:|:---|:---|:---|
"
    for e in emails:
        date_str = str(e.get('date', ''))[:16]
        sender = str(e.get('from', '')).replace('|', '&#124;')
        subject = str(e.get('subject', '')).replace('|', '&#124;')
        table += f"| {e['id']} | {date_str}.. | {sender} | {subject} |
"
    return table

@mcp.tool()
def list_emails(count: int = 10):
    """최근 이메일 목록(ID, 날짜, 발송인, 제목)만 조회합니다."""
    client, error = _get_client()
    if error: return error
    try:
        emails = client.list_emails(count)
        return _format_email_list_table(emails)
    except Exception as e:
        return f"❌ 목록 조회 실패: {str(e)}"

@mcp.tool()
def search_emails(keyword: str, limit: int = 50):
    """제목이나 보낸 사람 이름에서 키워드를 검색합니다."""
    client, error = _get_client()
    if error: return error
    try:
        results = client.search_emails(keyword, limit)
        if not results: return f"🔍 '{keyword}' 검색 결과가 없습니다."
        return f"🔍 **'{keyword}' 검색 결과**\n\n" + _format_email_list_table(results)
    except Exception as e:
        return f"❌ 검색 실패: {str(e)}"

@mcp.tool()
def read_email(email_id: int):
    """특정 ID의 이메일 본문 내용을 읽습니다. 사용자의 명시적 요청 시에만 사용하십시오."""
    client, error = _get_client()
    if error: return error
    try:
        email_data = client.get_email(email_id)
        if not email_data: return f"❌ ID {email_id}번 이메일을 찾을 수 없습니다."
        att_str = ", ".join([f'`{f}`' for f in email_data.get('attachments', [])]) or "(없음)"
        return f"\n# 📧 이메일 상세 내용 (ID: {email_data['id']})\n\n| 항목 | 내용 |\n|---|---|
| **보낸 사람** | {email_data['from']} |
| **날짜** | {email_data['date']} |
| **제목** | {email_data['subject']} |
| **첨부파일** | {att_str} |

## 📝 본문 내용
---
{email_data['body']}
---"
    except Exception as e:
        return f"❌ 메일 읽기 실패: {str(e)}"

@mcp.tool()
def send_email(to_email: str, subject: str, body: str):
    """이메일을 발송합니다. 사용자의 명시적 승인이 있을 때만 실행하십시오."""
    client, error = _get_client()
    if error: return error
    try:
        return client.send_email(to_email, subject, body)
    except Exception as e:
        return f"❌ 이메일 발송 중 에러 발생: {str(e)}"

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """이메일의 첨부파일을 다운로드합니다."""
    client, error = _get_client()
    if error: return error
    try:
        result = client.download_file(email_id, filename, save_path)
        if result.startswith("Success"):
            return f"✅ **다운로드 완료!**\n- 결과: `{result.replace('Success: ', '')}`"
        return f"❌ **다운로드 실패**: {result}"
    except Exception as e:
        return f"❌ 에러 발생: {str(e)}"

@mcp.tool()
def read_document(file_path: str):
    """로컬 문서(.pptx, .xlsx, .pdf 등)의 내용을 텍스트로 읽습니다."""
    if not os.path.exists(file_path): return f"❌ 파일이 존재하지 않습니다: `{file_path}`"
    try:
        content = document_loader.extract_text_from_file(file_path)
        return f"## 📄 문서 내용: {os.path.basename(file_path)}\n\n```text\n{content}\n```"
    except Exception as e:
        return f"❌ 문서 읽기 실패: {str(e)}"

def main():
    mcp.run()

if __name__ == "__main__":
    main()