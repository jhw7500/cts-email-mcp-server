import json
import os
import sys
from mcp.server.fastmcp import FastMCP

def log(msg):
    print(f"[*] {msg}", file=sys.stderr)

try:
    from .client import EmailClient
    from . import document_loader
except ImportError:
    try:
        from client import EmailClient
        import document_loader
    except ImportError:
        log("Critical: Could not load internal modules")

mcp = FastMCP("Cantops-Email")

def _get_client():
    try:
        user = os.environ.get("EMAIL_USER")
        pw = os.environ.get("EMAIL_PASSWORD")
        if not user or not pw:
            return None, "❌ 설정 오류: EMAIL_USER, EMAIL_PASSWORD 환경 변수를 확인해주세요."
        return EmailClient(), None
    except Exception as e:
        return None, f"❌ 클라이언트 초기화 에러: {str(e)}"

def _format_email_list_table(emails):
    if not emails:
        return "📭 조회된 이메일이 없습니다."
    table = "| ID | 날짜 | 보낸 사람 | 제목 |\n|---:|:---|:---|:---|"
    for e in emails:
        date_str = str(e.get('date', ''))[:16]
        sender = str(e.get('from', '')).replace('|', '&#124;')
        subject = str(e.get('subject', '')).replace('|', '&#124;')
        table += f"| {e['id']} | {date_str}.. | {sender} | {subject} |\n"
    return table

@mcp.tool()
def list_emails(count: int = 10):
    """
    [지침] 최근 이메일 목록(ID, 날짜, 발송인, 제목)만 조회합니다.
    - 사용자가 '메일 확인', '조회' 등 일반적인 요청을 할 때 이 도구만 사용하십시오.
    - 목록을 보여준 후, 사용자의 추가 명령 없이 자동으로 read_email이나 download_attachment를 호출하지 마십시오.
    """
    client, error = _get_client()
    if error: return error
    try:
        emails = client.list_emails(count)
        return _format_email_list_table(emails)
    except Exception as e:
        return f"❌ 목록 조회 실패: {str(e)}"

@mcp.tool()
def search_emails(keyword: str, limit: int = 50):
    """
    [지침] 제목이나 보낸 사람 이름에서 키워드를 검색합니다.
    - 검색 결과만 깔끔한 표 형식으로 보여주십시오.
    - 검색 후 자동으로 다른 도구를 호출하지 마십시오.
    """
    client, error = _get_client()
    if error: return error
    try:
        results = client.search_emails(keyword, limit)
        if not results:
            return f"🔍 '{keyword}'에 대한 검색 결과가 없습니다."
        return f"🔍 **'{keyword}' 검색 결과**\n\n" + _format_email_list_table(results)
    except Exception as e:
        return f"❌ 검색 실패: {str(e)}"

@mcp.tool()
def read_email(email_id: int):
    """
    [지침] 특정 ID의 이메일 본문 내용을 읽습니다.
    - 사용자가 특정 메일 ID를 지정하여 '읽어줘' 또는 '내용 보여줘'라고 명시적으로 요청한 경우에만 호출하십시오.
    - 목록 조회(list_emails) 결과에 따라 AI가 판단하여 자동으로 이 도구를 호출하는 것을 엄격히 금지합니다.
    """
    client, error = _get_client()
    if error: return error
    try:
        email_data = client.get_email(email_id)
        if not email_data: return f"❌ ID {email_id}번 이메일을 찾을 수 없습니다."
        att_str = ", ".join([f'`{f}`' for f in email_data.get('attachments', [])]) or "(없음)"
        return f"\n# 📧 이메일 상세 내용 (ID: {email_data['id']})\n\n| 항목 | 내용 |\n|---|---|
| **보낸 사람** | {email_data['from']} |\n| **날짜** | {email_data['date']} |\n| **제목** | {email_data['subject']} |\n| **첨부파일** | {att_str} |\n\n## 📝 본문 내용\n---\n{email_data['body']}\n---"
    except Exception as e:
        return f"❌ 메일 읽기 실패: {str(e)}"

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """
    [지침] 특정 이메일의 첨부파일을 다운로드합니다.
    - 사용자가 파일명을 언급하며 '다운로드해줘', '파일 저장해줘'라고 명시적으로 요청한 경우에만 호출하십시오.
    - AI가 필요하다고 판단하여 자동으로 파일을 다운로드하는 행위를 엄격히 금지합니다.
    """
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
    """
    [지침] 로컬 파일(.pptx, .xlsx, .pdf, .txt 등)의 내용을 텍스트로 읽습니다.
    - 사용자가 특정 파일을 읽어달라고 요청했을 때만 사용하십시오.
    """
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