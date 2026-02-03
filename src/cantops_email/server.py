import json
import os
import sys
import re
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
    [지침] 특정 ID의 이메일 본문 내용을 읽습니다. 명시적 요청 시에만 사용하십시오.
    """
    client, error = _get_client()
    if error: return error
    try:
        email_data = client.get_email(email_id)
        if not email_data: return f"❌ ID {email_id}번 이메일을 찾을 수 없습니다."
        att_str = ", ".join([f'`{f}`' for f in email_data.get('attachments', [])]) or "(없음)"
        return f"\n# 📧 이메일 상세 내용 (ID: {email_data['id']})\n\n| 항목 | 내용 |\n|---|---|
| **보낸 사람** | {email_data['from']} |\n| **날짜** | {email_data['date']} |\n| **제목** | {email_data['subject']} |\n| **첨부파일** | {att_str} |\n\n## 📝 본문 내용
---
{email_data['body']}
---"
    except Exception as e:
        return f"❌ 메일 읽기 실패: {str(e)}"

@mcp.tool()
def send_email(to_email: str, subject: str, body: str):
    """
    [🚨 중요: 지침] 이메일을 발송합니다.
    - 사용자가 발송 내용(수신인, 제목, 본문)을 확인하고 명시적으로 '보내줘'라고 승인한 경우에만 실행하십시오.
    - AI가 판단하여 자동으로 이메일을 발송하는 행위를 엄격히 금지합니다.
    - 테스트 시에는 수신인을 본인(환경변수의 EMAIL_USER)으로 설정하는 것을 권장합니다.
    """
    client, error = _get_client()
    if error: return error
    try:
        return client.send_email(to_email, subject, body)
    except Exception as e:
        return f"❌ 이메일 발송 중 에러 발생: {str(e)}"

@mcp.tool()
def analyze_and_extract_schedule(text: str):
    """
    [지침] 텍스트(이메일 본문 등)에서 날짜, 시간, 장소, 안건 등의 일정을 추출하여 정리합니다.
    - 이 도구는 실제 캘린더에 등록하지 않고, 등록하기 위한 정보를 구조화하여 제안만 합니다.
    """
    # 간단한 정규식 기반 추출 예시 (AI가 실제로는 텍스트를 분석하여 더 잘 정리함)
    date_pattern = r'\d{1,2}월\s?\d{1,2}일'
    time_pattern = r'\d{1,2}시(?:\s?\d{1,2}분)?'
    
    dates = re.findall(date_pattern, text)
    times = re.findall(time_pattern, text)
    
    report = "🗓️ **추출된 일정 제안**\n\n"
    if dates: report += f"- **날짜**: {', '.join(set(dates))}\n"
    if times: report += f"- **시간**: {', '.join(set(times))}\n"
    report += "\n---\n위 내용을 캘린더에 등록하시겠습니까? (현재는 정보 추출 기능만 제공됩니다.)"
    return report

@mcp.tool()
def download_attachment(email_id: int, filename: str, save_path: str = "./downloads"):
    """
    [지침] 특정 이메일의 첨부파일을 다운로드합니다. 명시적 요청 시에만 사용하십시오.
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
    [지침] 로컬 파일의 내용을 텍스트로 읽습니다.
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
