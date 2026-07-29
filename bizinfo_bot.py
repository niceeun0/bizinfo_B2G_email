# -*- coding: utf-8 -*-
import os
import re
import urllib.request
import json
import ssl
import smtplib
import io
import time
import requests
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openai import OpenAI
from pypdf import PdfReader

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

# OpenRouter 설정 (GEMINI_API_KEY 환경변수 활용)
API_KEY = os.environ.get("GEMINI_API_KEY")
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=API_KEY,
) if API_KEY else None

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")

def fetch_bizinfo_notices():
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=50"
    
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
            print(f"⏳ [수집 시도 {attempt}/{max_retries}] 표준 보안 연결로 기업마당 API 호출 중...")
            with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
                if response.status == 200:
                    res_body = response.read().decode('utf-8')
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    print(f"🎯 [수집 성공] 총 {len(items)}건의 공고를 불러왔습니다.")
                    return items
        except Exception as e:
            print(f"⚠️ [연결 경고 (시도 {attempt})]: {str(e)}")
            if attempt < max_retries:
                time.sleep(3)
            else:
                print("❌ [최종 실패] API 연결에 실패했습니다.")
    return []

def extract_text_from_pdf_url(pdf_url):
    if not pdf_url or not pdf_url.startswith("http"):
        return ""
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        res = requests.get(pdf_url, headers=headers, timeout=10)
        if res.status_code == 200 and len(res.content) > 1000:
            if not res.content.startswith(b'%PDF'):
                return ""
            pdf_file = io.BytesIO(res.content)
            reader = PdfReader(pdf_file)
            extracted_text = ""
            for page in reader.pages[:8]:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
            return extracted_text.strip()
    except Exception:
        pass
    return ""

def extract_relevant_snippets(text):
    """토큰 절약을 위해 핵심 키워드 주변 문맥만 추출합니다."""
    if not text:
        return ""
    
    keywords = ["규모", "선정", "지원", "개사", "팀", "서류", "제출", "신청서", "접수", "이메일", "담당", "문의", "사업자", "등기", "재무제표", "납세", "완납", "계획서"]
    lines = text.split('\n')
    relevant_lines = []
    
    for line in lines:
        if any(kw in line for kw in keywords):
            relevant_lines.append(line.strip())
            
    snippets = "\n".join(relevant_lines)
    if len(snippets) > 2500:
        snippets = snippets[:2500]
    return snippets

def analyze_with_openrouter(full_text):
    """필수 제출 서류를 누락 없이 구체적으로 추출하도록 프롬프트를 강화합니다."""
    scale_result = "지원 규모 확인 필요"
    contact_result = "담당자 정보 확인 필요"
    docs_result = "필수 서류 확인 필요"

    snippets = extract_relevant_snippets(full_text)
    if not client or not snippets:
        return scale_result, contact_result, docs_result
    
    try:
        completion = client.chat.completions.create(
            model="openrouter/free",
            messages=[
                {
                    "role": "system",
                    "content": "당신은 정부 지원사업 공고 분석 전문가입니다. 불필요한 설명이나 영어를 절대 출력하지 말고, 오직 지정된 3가지 양식만 정확히 한국어로 답변하세요."
                },
                {
                    "role": "user",
                    "content": f"""
다음 공고 핵심 내용을 분석하여 정확히 아래 3줄 형식으로만 답변해 주세요.

1. 지원규모: 선정하는 기업 수와 지원 내용/금액을 간결하게 요약 (예: 5개사 내외 선정, 각 최대 5천만원 지원 등). 정보가 없으면 '확인 필요'
2. 담당접수처: 담당자 성명, 전화번호, 접수 이메일 주소 등 연락처 정보. 정보가 없으면 '확인 필요'
3. 필수서류: 정부 지원사업에서 요구하는 필수 제출 서류들(예: 사업자등록증, 법인등기부등본, 표준재무제표증명, 부가가치세과세표준증명원, 국세/지방세 납세증명서, 사업계획서, 개인정보동의서 등)을 공고문 내용에 근거하여 빠짐없이 구체적으로 쉼표로 나열하세요. 정보가 없으면 '확인 필요'

반드시 아래 양식 그대로 3줄만 작성하세요. 다른 사족은 절대 금지합니다.

지원규모: [내용 작성]
담당접수처: [내용 작성]
필수서류: [내용 작성]

[핵심 공고 내용]
{snippets}
                    """
                }
            ],
            temperature=0.1,
            max_tokens=300,
        )
        
        if completion and completion.choices and completion.choices[0].message:
            text_resp = completion.choices[0].message.content
            if text_resp:
                text_resp = text_resp.strip()
                for line in text_resp.split('\n'):
                    if "지원규모:" in line or "모집규모:" in line:
                        scale_result = line.split(":", 1)[1].strip()
                    elif "담당접수처:" in line or "문의처:" in line:
                        contact_result = line.split(":", 1)[1].strip()
                    elif "필수서류:" in line:
                        docs_result = line.split(":", 1)[1].strip()
                        
    except Exception as e:
        print(f"⚠️ OpenRouter API 분석 중 오류: {str(e)}")
        
    return scale_result, contact_result, docs_result

def analyze_and_build_html(items):
    rows_html = ""
    valid_count = 0
    seen_signatures = set()
    
    for idx, item in enumerate(items, 1):
        title = item.get("pblancNm") or item.get("title", "제목없음")
        exec_org = item.get("excInsttNm") or item.get("jrsdInsttNm") or "수행기관 미표기"
        period = item.get("reqstBeginEndDe") or item.get("reqstDt", "일정 참조")
        link = item.get("pblancUrl") or item.get("link", "#")
        summary = item.get("bsnsSumryCn", "")
        ref_name = item.get("refrncNm", "")
        original_method = item.get("reqstMthDscd") or item.get("reqstMthCn") or item.get("reqstMthPapersCn") or item.get("rcivMth") or "공고문 참조"
        
        method_lower = original_method.lower()
        is_online_only = ("온라인" in original_method or "누리집" in original_method or "홈페이지" in original_method) and not ("이메일" in original_method or "전자우편" in original_method or "방문" in original_method or "우편" in original_method)
        
        pdf_url = item.get("printFlpthNm", "")
        pdf_text = extract_text_from_pdf_url(pdf_url)
        full_text = f"{summary} {ref_name} {original_method} {pdf_text}"

        raw_emails = EMAIL_PATTERN.findall(full_text)
        emails = []
        for e in raw_emails:
            if "@" in e and "." in e:
                emails.append(e)
        emails = list(set(emails))
        
        is_valid_apply = (len(emails) > 0) or ("이메일" in original_method) or ("전자우편" in original_method) or ("방문" in original_method) or ("우편" in original_method)
        
        if is_online_only or not is_valid_apply:
            continue

        base_title = re.sub(r'\[.*?\]', '', title).strip()
        base_title = re.sub(r'(맞춤형|마케팅|디자인|기술혁신)', '', base_title).strip()
        rep_email = emails[0] if emails else ""
        
        signature = f"{exec_org}_{rep_email}_{base_title[:15]}"
        
        if signature in seen_signatures:
            print(f"🔄 [중복 공고 스킵]: '{title}'")
            continue
        seen_signatures.add(signature)

        valid_count += 1
        print(f"🤖 [OpenRouter AI 토큰 최적화 단건 분석 ({valid_count})] '{title}'...")
        
        parsed_scale, parsed_contact, parsed_docs = analyze_with_openrouter(full_text)
        time.sleep(0.3)

        target_type = "👤 개인" if ("개인" in full_text and "기업" not in full_text) else "🏢 기업"
        email_str = ", ".join(emails) if emails else ""
        email_html = "<br>".join([f"📧 {e}" for e in emails]) if emails else "-"
        
        if len(emails) > 0:
            to_emails = emails[0]
            cc_emails = ",".join(emails[1:]) if len(emails) > 1 else ""
            mailto_link = f"mailto:{to_emails}"
            if cc_emails:
                mailto_link += f"?cc={cc_emails}"
            one_click_html = f'<a href="{mailto_link}" style="background:#e6f4ea; color:#137333; padding:6px 12px; border-radius:4px; font-size:11px; font-weight:bold; text-decoration:none; display:inline-block;">⚡ 원클릭 메일 제안</a>'
        else:
            one_click_html = '<span style="color:#137333; font-weight:bold; font-size:11px;">접수 공고</span>'

        scale_cell = f'<span style="color:#1a73e8; font-weight:bold; font-size:12px;">{parsed_scale}</span>'
        contact_cell = f'<span style="color:#333; font-size:11px;">{parsed_contact}</span>'
        docs_cell = f'<span style="color:#c5221f; font-weight:bold; font-size:11px;">{parsed_docs}</span>'
        
        phones = PHONE_PATTERN.findall(full_text)
        fallback_phone = f"📞 {phones[0]}" if phones else (ref_name if ref_name else "-")

        safe_title = title.replace('"', '""').replace('\n', ' ')
        safe_org = exec_org.replace('"', '""')
        safe_scale = parsed_scale.replace('"', '""')
        safe_contact = parsed_contact.replace('"', '""')
        safe_docs = parsed_docs.replace('"', '""')
        safe_emails = email_str.replace('"', '""')

        rows_html += f"""
        <tr data-row='{{"구분":"{target_type}","공고명":"{safe_title}","지원규모":"{safe_scale}","수행기관":"{safe_org}","신청기간":"{period}","신청방법":"{original_method}","담당접수처":"{safe_contact}","이메일":"{safe_emails}","필수서류":"{safe_docs}"}}'>
            <td style="padding:12px; border-bottom:1px solid #eee; text-align:center; font-size:11px;">{target_type}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; font-weight:bold;">{title}</td>
            <td style="padding:12px; border-bottom:1px solid #eee;">{scale_cell}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; color:#555;">{exec_org}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; font-size:12px;">{period}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; font-size:12px;">{original_method}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; text-align:center;">{one_click_html}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; font-size:12px;">{contact_cell}<br>{email_html}</td>
            <td style="padding:12px; border-bottom:1px solid #eee;">{docs_cell}</td>
            <td style="padding:12px; border-bottom:1px solid #eee; text-align:center;"><a href="{link}" target="_blank" style="background:#004aad; color:white; padding:6px 10px; text-decoration:none; border-radius:4px; font-size:11px;">보기</a></td>
        </tr>
        """

    if valid_count == 0:
        rows_html = """
        <tr>
            <td colspan="10" style="padding:20px; text-align:center; color:#666;">조건에 맞는 유효한 접수 공고가 없습니다.</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <script>
            function downloadCSV() {{
                let rows = document.querySelectorAll("table tbody tr[data-row]");
                if (rows.length === 0) {{
                    alert("다운로드할 데이터가 없습니다.");
                    return;
                }}
                let csvContent = "\\uFEFF";
                csvContent += "구분,공고명,지원규모,수행기관,신청기간,신청방법,담당접수처,이메일,필수서류\\n";
                
                rows.forEach(function(row) {{
                    let data = JSON.parse(row.getAttribute("data-row"));
                    let line = [
                        '"' + (data.구분 || "") + '"',
                        '"' + (data.공고명 || "") + '"',
                        '"' + (data.지원규모 || "") + '"',
                        '"' + (data.수행기관 || "") + '"',
                        '"' + (data.신청기간 || "") + '"',
                        '"' + (data.신청방법 || "") + '"',
                        '"' + (data.담당접수처 || "") + '"',
                        '"' + (data.이메일 || "") + '"',
                        '"' + (data.필수서류 || "") + '"'
                    ].join(",");
                    csvContent += line + "\\n";
                }});
                
                let blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
                let url = URL.createObjectURL(blob);
                let a = document.createElement("a");
                a.href = url;
                a.download = "B2G_AI_Optimized_Report.csv";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }}
        </script>
    </head>
    <body style="font-family:'Malgun Gothic', sans-serif; color:#333;">
        <div style="max-width:1700px; margin:0 auto; padding:20px;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <h2 style="color:#004aad; margin:0;">📊 B2G 알짜배기 공고 & OpenRouter AI 토큰 최적화 대시보드</h2>
                <button onclick="downloadCSV()" style="background:#137333; color:white; border:none; padding:10px 16px; border-radius:6px; font-weight:bold; cursor:pointer; font-size:13px;">📥 대시보드 엑셀(CSV) 다운로드</button>
            </div>
            <p style="margin-top:10px;">토큰 최적화 및 상세 서류 추출 필터링을 거친 이메일/오프라인 접수 가능 공고 총 <b>{valid_count}건</b>의 AI 요약 리포트입니다.</p>
            <table style="width:100%; border-collapse:collapse; margin-top:15px; background:#fff;">
                <thead>
                    <tr style="background-color:#f8fafc; color:#444; font-size:12px;">
                        <th style="padding:12px; border-bottom:2px solid #ddd;">구분</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd; width:20%;">공고명</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd; width:15%;">📌 선정/지원규모</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd;">사업수행기관</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd;">신청기간</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd;">신청방법</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd;">원클릭 메일 제안</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd; width:15%;">📞 담당접수처 및 이메일</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd; width:20%;">📋 AI 분석 필수 서류</th>
                        <th style="padding:12px; border-bottom:2px solid #ddd;">바로가기</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html_content

def send_email(html_body):
    sender_email = os.environ.get("MAIL_USER")
    sender_password = os.environ.get("EMAIL_PASS")
    receiver_list = os.environ.get("MAIL_RECEIVER") or sender_email

    if not sender_email or not sender_password:
        print("⚠️ [이메일 생략] 깃허브 Secrets에 MAIL_USER 또는 EMAIL_PASS가 설정되지 않았습니다.")
        return

    receivers = [email.strip() for email in receiver_list.split(",")]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "📊 [B2G 토큰 최적화 AI] 공고별 선정규모·담당자·필수서류 상세 분석 결과"
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
        html_report = analyze_and_build_html(notices)
        send_email(html_report)
    else:
        print("❌ 공고 데이터를 가져오지 못해 작업을 종료합니다.")
