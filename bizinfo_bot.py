# -*- coding: utf-8 -*-
import os
import re
import urllib.request
import json
import ssl
import smtplib
import io
import time
import tempfile
import zipfile
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openai import OpenAI
import urllib3
import base64

# HTTPS 보안 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

# OpenRouter 설정
API_KEY = os.environ.get("GEMINI_API_KEY")
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=API_KEY,
) if API_KEY else None

# 문서 파싱 및 OCR 라이브러리 안전 로드
try: import pdfplumber
except ImportError: pdfplumber = None
try: import docx
except ImportError: docx = None
try: import olefile
except ImportError: olefile = None
try: import pytesseract
except ImportError: pytesseract = None

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")
DOC_KEYWORDS = ["사업자등록증", "재무제표", "인감증명서", "주주명부", "법인등기부등본", "신용평가등급", "신용평가서", "기업정보"]
SUBMIT_KEYWORDS = ["우편", "이메일", "e-mail", "메일", "방문", "직접", "사본", "스캔본"]

def fetch_bizinfo_notices():
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=100"
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        target_url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            "Accept": "application/json"
        }
    )

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            print(f"⏳ [수집 시도 {attempt}/{max_retries}] 기업마당 API 호출 중...")
            with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
                if response.status == 200:
                    res_body = response.read().decode('utf-8')
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    print(f"🎯 [수집 성공] 총 {len(items)}건의 공고를 불러왔습니다.")
                    return items
        except Exception as e:
            print(f"⚠️ [연결 경고 (시도 {attempt})]: {str(e)}")
            if attempt < max_retries: time.sleep(3)
    return []

def extract_docs_and_contacts(urls: list) -> tuple:
    if not urls: return "", "", "-", "첨부파일 없음"
    emails, phones = set(), set()
    needs_oneclick = False
    
    for idx, url in enumerate(urls[:2]):
        try:
            resp = requests.get(url, timeout=30, verify=False)
            if resp.status_code != 200: continue
            
            content = resp.content
            text = ""
            
            # PDF 판독 및 OCR
            if content.startswith(b'%PDF'):
                if pdfplumber:
                    try:
                        with pdfplumber.open(io.BytesIO(content)) as pdf:
                            pages_to_scan = pdf.pages[:3] + pdf.pages[-3:] if len(pdf.pages) > 6 else pdf.pages
                            for page in pages_to_scan:
                                t = page.extract_text()
                                if t: text += t + "\n"
                                if pytesseract and (not t or len(t.strip()) < 30):
                                    try:
                                        img = page.to_image(resolution=150).original
                                        text += pytesseract.image_to_string(img, lang='eng+kor') + "\n"
                                    except Exception: pass
                    except Exception: pass
            # DOCX / HWPX 판독
            elif content.startswith(b'PK'):
                parsed = False
                if docx:
                    try:
                        doc = docx.Document(io.BytesIO(content))
                        text = "\n".join([p.text for p in doc.paragraphs])
                        parsed = True
                    except Exception: pass
                if not parsed:
                    try:
                        with zipfile.ZipFile(io.BytesIO(content)) as zf:
                            for name in zf.namelist():
                                if "contents" in name and name.endswith(".xml"):
                                    xml_content = zf.read(name).decode("utf-8", errors="ignore")
                                    text += re.sub("<[^>]+>", " ", xml_content) + "\n"
                    except Exception: pass
            # HWP 판독
            elif content.startswith(b'\xd0\xcf\x11\xe0'):
                if olefile:
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".hwp") as tmp:
                            tmp.write(content)
                            tmp_path = tmp.name
                        if olefile.isOleFile(tmp_path):
                            ole = olefile.OleFileIO(tmp_path)
                            if ole.exists("PrvText"):
                                text = ole.openstream("PrvText").read().decode("utf-16", errors="ignore")
                            ole.close()
                        os.remove(tmp_path)
                    except Exception: pass
            
            emails.update(EMAIL_PATTERN.findall(text))
            phones.update(PHONE_PATTERN.findall(text))
            
            if not needs_oneclick:
                if any(kw in text for kw in DOC_KEYWORDS) and any(kw in text for kw in SUBMIT_KEYWORDS):
                    needs_oneclick = True
                    
        except Exception:
            pass
            
    oneclick_flag = "💡 원클릭 제안 필요" if needs_oneclick else "-"
    return ", ".join(sorted(emails)), ", ".join(sorted(phones)), oneclick_flag

def clean_mixed_text(text):
    if not text: return "확인 필요"
    lines = text.split('\n')
    cleaned_lines = [l for l in lines if not (sum(1 for c in l if ord(c) < 128 and c.isalpha()) / max(len(l.strip()), 1) > 0.6)]
    result = " ".join(cleaned_lines).strip()
    if len(result) < 2:
        result = "".join(re.findall(r'[가-힣0-9%.,~()\-\s]+', text)).strip()
    return result if len(result) > 1 else "확인 필요"

def analyze_with_openrouter(full_text):
    scale_result, docs_result = "지원 규모 확인 필요", "필수 서류 확인 필요"
    if not client or not full_text: return scale_result, docs_result
    
    try:
        completion = client.chat.completions.create(
            model="openrouter/free",
            messages=[
                {"role": "system", "content": "당신은 정부 지원사업 공고 분석 전문가입니다. 반드시 순수 한국어로만 핵심만 요약하십시오."},
                {"role": "user", "content": f"다음 내용을 분석하여 정확히 두 줄로 요약해 주세요.\n\n지원규모: [선정 기업 수 및 지원 내용]\n필수서류: [제출 서류 나열]\n\n{full_text[:3000]}"}
            ],
            temperature=0.1,
            max_tokens=250,
        )
        if completion and completion.choices and completion.choices[0].message:
            text_resp = completion.choices[0].message.content.strip()
            for line in text_resp.split('\n'):
                if "지원규모:" in line: scale_result = clean_mixed_text(line.split(":", 1)[1])
                elif "필수서류:" in line: docs_result = clean_mixed_text(line.split(":", 1)[1])
    except Exception: pass
    return scale_result, docs_result

def analyze_and_build_newsletter(items):
    cards_html = ""
    valid_count = 0
    email_action_count = 0
    online_count = 0
    seen_signatures = set()
    csv_rows = []
    
    yesterday = (datetime.utcnow() + timedelta(hours=9) - timedelta(days=1)).strftime('%Y%m%d')
    yesterday_alt = yesterday[:4] + "-" + yesterday[4:6] + "-" + yesterday[6:]
    print(f"📅 [타겟 수집일 (전날 등록 공고 기준)]: {yesterday_alt}")

    for idx, item in enumerate(items, 1):
        reg_date = str(item.get("pblancDe") or item.get("regDt") or item.get("creatDt") or "")
        if reg_date:
            if yesterday not in reg_date and yesterday_alt not in reg_date: continue
        else:
            if idx > 20: continue

        title = item.get("pblancNm") or item.get("title", "제목없음")
        exec_org = item.get("excInsttNm") or item.get("jrsdInsttNm") or "수행기관 미표기"
        period = item.get("reqstBeginEndDe") or item.get("reqstDt", "일정 참조")
        link = item.get("pblancUrl") or item.get("link", "#")
        summary = item.get("bsnsSumryCn", "")
        original_method = item.get("reqstMthDscd") or item.get("reqstMthCn") or item.get("rcivMth") or "공고문 참조"
        
        attach_urls = []
        for f_key in ["bidSpecificationUrl"] + [f"ntceSpecDocUrl{i}" for i in range(1, 6)]:
            u = item.get(f_key)
            if u:
                if not u.startswith("http"): u = "https://www.g2b.go.kr:8081" + ("/" if not u.startswith("/") else "") + u
                attach_urls.append(u)

        doc_email, doc_tel, oneclick_flag = "", "", "-"
        if attach_urls:
            doc_email, doc_tel, oneclick_flag = extract_docs_and_contacts(attach_urls)

        full_text = f"{title} {summary} {original_method}"

        raw_emails = EMAIL_PATTERN.findall(full_text)
        if doc_email: raw_emails.extend(doc_email.split(", "))
        emails = list(set([e.strip() for e in raw_emails if "@" in e and len(e) < 50]))
        
        clean_title_sig = re.sub(r'\s+', '', title)
        signature = f"{exec_org}_{clean_title_sig}"
        if signature in seen_signatures: continue
        seen_signatures.add(signature)

        is_online_only = ("온라인" in original_method or "누리집" in original_method or "홈페이지" in original_method) and not ("이메일" in original_method or "전자우편" in original_method or "방문" in original_method or "우편" in original_method)
        
        if is_online_only:
            online_count += 1
            badge_text, badge_style, target_type = "온라인 사이트 접수", "background:#e0f2fe; color:#0369a1;", "온라인"
        else:
            email_action_count += 1
            badge_text, badge_style, target_type = "이메일/방문접수", "background:#e6f4ea; color:#137333;", "직접접수"

        valid_count += 1
        print(f"🤖 [AI 분석 중 ({valid_count})] '{title}'...")
        parsed_scale, parsed_docs = analyze_with_openrouter(full_text)
        time.sleep(0.2)

        phones = PHONE_PATTERN.findall(full_text)
        if doc_tel: phones.extend(doc_tel.split(", "))
        phone_str = phones[0] if phones else "문의처 참조"
        email_str = emails[0] if emails else "이메일 공고문 참조"

        if not is_online_only and emails:
            mailto_link = f"mailto:{emails[0]}?subject=[지원문의] {title}"
            action_btn = f'<a href="{mailto_link}" style="background:#137333; color:white; padding:6px 14px; border-radius:4px; font-size:11px; font-weight:bold; text-decoration:none; display:inline-block;">⚡ 원클릭 메일 제안</a>'
        else:
            action_btn = f'<a href="{link}" target="_blank" style="background:#0369a1; color:white; padding:6px 12px; border-radius:4px; font-size:11px; font-weight:bold; text-decoration:none;">🌐 온라인 신청 바로가기</a>'

        safe_title = title.replace('"', '').replace("'", "").replace('\n', ' ')
        safe_org = exec_org.replace('"', '').replace("'", "")
        safe_scale = parsed_scale.replace('"', '').replace("'", "").replace('\n', ' ')
        safe_docs = parsed_docs.replace('"', '').replace("'", "").replace('\n', ' ')
        safe_emails = email_str.replace('"', '').replace("'", "")
        safe_method = original_method.replace('"', '').replace("'", "").replace('\n', ' ')

        csv_rows.append([target_type, safe_title, safe_scale, safe_org, period, safe_method, safe_emails, safe_docs])

        cards_html += f"""
        <div style="background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin-bottom:15px; padding:18px 20px;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px;">
                <a href="{link}" target="_blank" style="font-size:15px; font-weight:bold; color:#1e293b; text-decoration:none; line-height:1.4; flex:1;">{title}</a>
                <span style="{badge_style} padding:3px 8px; border-radius:4px; font-size:11px; font-weight:bold; white-space:nowrap; margin-left:10px;">{badge_text}</span>
            </div>
            <div style="display:grid; grid-template-columns: repeat(2, 1fr); gap:8px; font-size:12px; color:#64748b; margin-bottom:12px;">
                <div>🏢 <b>수행기관:</b> {exec_org}</div>
                <div>📞 <b>담당연락처:</b> {phone_str}</div>
                <div>📧 <b>접수메일:</b> {email_str}</div>
                <div>⏳ <b>신청방법:</b> {original_method[:30]}</div>
            </div>
            <div style="background:#f8fafc; border-left:3px solid #004aad; padding:10px 12px; border-radius:4px; font-size:12px; margin-bottom:12px;">
                <p style="margin:3px 0;"><span style="font-weight:bold; color:#004aad;">📌 지원규모:</span> {parsed_scale}</p>
                <p style="margin:3px 0;"><span style="font-weight:bold; color:#004aad;">📋 필수서류:</span> {parsed_docs}</p>
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center; border-top:1px solid #f1f5f9; padding-top:12px;">
                <span style="font-size:11px; color:#94a3b8;">신청기간: {period}</span>
                <div style="display:flex; gap:8px;">
                    {action_btn}
                    <a href="{link}" target="_blank" style="background:#f1f5f9; color:#475569; padding:6px 12px; border-radius:4px; font-size:11px; font-weight:bold; text-decoration:none;">공고 원문보기</a>
                </div>
            </div>
        </div>
        """

    if valid_count == 0:
        cards_html = '<div style="background:#fff; border:1px solid #e2e8f0; border-radius:8px; padding:30px; text-align:center; color:#666;">전날 등록된 신규 공고가 없거나 조건에 맞는 공고가 없습니다.</div>'

    csv_text = "\uFEFF" + "구분,공고명,지원규모,수행기관,신청기간,신청방법,이메일,필수서류\n"
    for r in csv_rows:
        csv_text += f'"{r[0]}","{r[1]}","{r[2]}","{r[3]}","{r[4]}","{r[5]}","{r[6]}","{r[7]}"\n'
    b64_csv = base64.b64encode(csv_text.encode('utf-8-sig')).decode('utf-8')
    csv_data_uri = f"data:text/csv;charset=utf-8;base64,{b64_csv}"

    today_str = datetime.utcnow().strftime('%Y.%m.%d (%a)')
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family:'Malgun Gothic', sans-serif; background-color:#f4f6f9; margin:0; padding:20px; color:#333;">
        <div style="max-width:900px; margin:0 auto; background:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 4px 15px rgba(0,0,0,0.05);">
            <div style="background:#004aad; color:white; padding:25px 30px; display:flex; justify-content:space-between; align-items:center;">
                <div>
                    <h2 style="margin:0; font-size:20px;">📊 B2G 알짜배기 공고 뉴스레터</h2>
                    <p style="margin:5px 0 0 0; font-size:13px; opacity:0.9;">메일 발송 전날 등록된 신규 공고 및 제안 리포트</p>
                </div>
                <div style="text-align:right;">
                    <a href="{csv_data_uri}" download="B2G_AI_Report.csv" style="background:#137333; color:white; border:none; padding:8px 14px; border-radius:6px; font-weight:bold; cursor:pointer; font-size:12px; text-decoration:none; display:inline-block; margin-bottom:5px;">📥 엑셀(CSV) 다운로드</a>
                    <div style="font-size:11px; opacity:0.9;">{today_str} 기준 (총 {valid_count}건)</div>
                </div>
            </div>
            <div style="background:#f8fafc; padding:15px 30px; border-bottom:1px solid #eee; display:flex; gap:20px; font-size:13px; font-weight:bold; color:#555;">
                <span>📁 신규 수집 공고: {valid_count}건</span>
                <span style="color:#137333;">⚡ 원클릭 메일 제안 가능: {email_action_count}건</span>
                <span style="color:#0369a1;">🌐 온라인 접수: {online_count}건</span>
            </div>
            <div style="padding:20px 30px;">
                {cards_html}
            </div>
            <div style="background:#f8fafc; text-align:center; padding:15px; font-size:11px; color:#64748b; border-top:1px solid #eee;">
                본 메일은 B2G AI 분석 봇에 의해 자동으로 생성되었습니다.
            </div>
        </div>
    </body>
    </html>
    """
    return html_content

def send_email(html_body):
    sender_email = os.environ.get("MAIL_USER")
    sender_password = os.environ.get("EMAIL_PASS")
    receiver_list = os.environ.get("MAIL_RECEIVER") or sender_email
    if not sender_email or not sender_password: return

    receivers = [email.strip() for email in receiver_list.split(",")]
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "📊 [B2G AI 뉴스레터] 전날 등록된 신규 지원사업 공고 리포트"
    msg["From"] = sender_email
    msg["To"] = ", ".join(receivers)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receivers, msg.as_string())
        server.quit()
        print(f"🎉 [이메일 전송 성공] 수신자: {receivers}")
    except Exception as e:
        print(f"❌ [이메일 전송 실패]: {str(e)}")

if __name__ == "__main__":
    notices = fetch_bizinfo_notices()
    if notices:
        html_report = analyze_and_build_newsletter(notices)
        send_email(html_report)
