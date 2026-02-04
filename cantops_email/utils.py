import email.header
from bs4 import BeautifulSoup

def clean_header(value):
    if not value: return ""
    try:
        decoded = email.header.decode_header(value)
        result = ""
        for text, encoding in decoded:
            if isinstance(text, bytes):
                result += text.decode(encoding or "utf-8", errors="ignore")
            else:
                result += str(text)
        return result
    except:
        return str(value)

def get_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = part.get("Content-Disposition")
            if ctype == "text/plain" and (not cdispo or "attachment" not in cdispo):
                payload = part.get_payload(decode=True)
                if payload:
                    body += payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = payload.decode(msg.get_content_charset() or "utf-8", errors="ignore")
    
    if "<html>" in body.lower():
        try:
            return BeautifulSoup(body, "html.parser").get_text(separator="\n").strip()
        except: pass
    return body.strip()

def get_attachment_list(msg):
    files = []
    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            files.append(clean_header(filename))
    return files

