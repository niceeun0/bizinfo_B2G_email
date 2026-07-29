# -*- coding: utf-8 -*-
import io
import os
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

SENDER_EMAIL = "niceeun095@gmail.com"  
SENDER_PASSWORD = os.environ.get("EMAIL_PASS")

RECEIVER_EMAILS = [
    "niceeun095@gmail.com",
    "s_e_y_0615@naver.com",
    "eyson0615@nice.co.kr"
]

CORP_KEYWORDS = ["중소기업", "기업", "법인", "사업자", "컨소시엄", "주관기관", "벤처", "소상공인"]
INDIVIDUAL_KEYWORDS = ["개인", "일반국민", "청년", "구직자", "학생", "개인사업자 제외"]
DOC_KEYWORDS = ["사업자등록증", "재무제표", "신용평가", "인감증명서", "법인등기", "주주명부", "국세완납", "지방세완납"]

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")

def fetch_bizinfo_notices():
    req_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=10&pageIndex=1"
    
    headers = {
        "Host": "www.bizinfo.go.kr",
        "Connection": "keep-alive",
        "Cache-Control": "max-age=0",
        "DNT": "1",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.9,en;q=0.9"
    }
    
    try:
        print("⏳ 방화벽 우회 접속을 시도합니다...")
        session = requests.Session()
        resp = session.get(req_url, headers=headers, timeout=25, verify=False)
        print(f"✅ 서버 응답 상태 코드: {resp.status_code}")
        
        if resp.status_code == 200:
            try:
                data = resp.json()
                if "reqErr" in data:
                    print(f"❌ 서버 에러 메시지: {data['reqErr']}")
                    return []
                    
                items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                print(f"🎯 정상 파싱된 지원사업 공고 개수: {len(items)}개")
                return items
            except Exception as json_e:
                print(f"❌ JSON 파싱 에러: {json_e}")
                print(f"📦 원본 응답 내용 일부: {resp.text[:300]}")
                
    except Exception as e:
        print(f"❌ API 요청 중 에러 발생 (타임아웃 등): {e}")
        
    return []

def process_and_analyze(items):
    matched_results = []
    for item in items:
        title = item.get("pblancNm") or item.get("title", "")
        summary = item.get("bsnsSumryCn") or item.get("description", "")
        target_nm = item.get("trgetNm", "")
        org_name = item.get("jrsdInsttNm") or item.get("author", "-")
        exec_org = item.get("excInsttNm", "-")
        link = item.get("pblancUrl") or item.get("link", "#")
        file_url = item.get("flpthNm") or item.get("printFlpthNm", "")
        contact_info = item.get("refrncNm", "")
        papers_info = item.get("reqstMthPapersCn", "")
            
        target_text = f"{target_nm} {summary}"
        corp_score = sum(target_text.count(kw) for kw in CORP_KEYWORDS)
        ind_score = sum(target_text.count(kw) for kw in INDIVIDUAL_KEYWORDS)
        target_type = "🏢 기업 대상" if corp_score >= ind_score else "👤 개인 대상"
        
        extracted_docs = [doc for doc in DOC_KEYWORDS if doc in f"{summary} {papers_info}"]
        extracted_emails = set(EMAIL_PATTERN.findall(f"{summary} {contact_info}"))
        extracted_phones = set(PHONE_PATTERN.findall(f"{summary} {contact_info}"))
        
        if file_url and file_url.endswith(".pdf"):
            try:
                f_resp = requests.get(file_url, timeout=10, verify=False)
                if f_resp.status_code == 200 and f_resp.content.startswith(b'%PDF'):
                    with pdfplumber.open(io.BytesIO(f_resp.content)) as pdf:
                        pdf_text = "\n".join([p.extract_text() or "" for p in pdf.pages[:2]])
                        for doc in DOC_KEYWORDS:
                            if doc in pdf_text and doc not in extracted_docs:
                                extracted_docs.append(doc)
                        extracted_emails.update(EMAIL_PATTERN.findall(pdf_text))
                        extracted_phones.update(PHONE_PATTERN.findall(pdf_text))
            except Exception:
                pass
                
        matched_results.append({
            "공고명": title, "소관기관": org_name, "수행기관": exec_org,
            "지원대상": target_type, "신청기간": item.get("reqstBeginEndDe") or item.get("reqstDt", "-"),
            "필요서류": ", ".join(extracted_docs) if extracted_docs else "공고문 참조",
            "담당자이메일": ", ".join(sorted(extracted_emails)) or "-",
            "담당자연락처": ", ".join(sorted(extracted_phones)) or "-",
            "공고링크": link
        })
    return matched_results

def send_email_report(results):
    if not results:
        print("조회된 공고가 없습니다.")
        return

    today_str = datetime.now().strftime("%Y년 %m월 %d일")
    html_content = f"""
    <html>
    <body style="font-family: 'Malgun Gothic', sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #004aad;">📡 [기업마당] 오늘의 신규 지원사업 전체 리포트 ({today_str})</h2>
        <p>※ 최근 등록된 <b>{len(results)}건</b>의 전체 지원사업 공고입니다.</p>
        <hr style="border: 1px solid #ddd;">
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; text-align: left; font-size: 14px;">
            <tr style="background-color: #f2f4f7;">
                <th style="width: 12%;">구분</th><th style="width: 28%;">공고명 / 소관기관</th>
                <th style="width: 15%;">신청기간</th><th style="width: 20%;">필요 제출 서류</th>
                <th style="width: 15%;">담당자 연락처</th><th style="width: 10%;">바로가기</th>
            </tr>
    """
    for r in results:
        html_content += f"""
        <tr>
            <td><b>{r['지원대상']}</b></td>
            <td><b>{r['공고명']}</b><br><span style="font-size: 12px; color: #666;">기관: {r['소관기관']} ({r['수행기관']})</span></td>
            <td>{r['신청기간']}</td>
            <td><span style="color: #d93025; font-weight: bold;">{r['필요서류']}</span></td>
            <td>📧 {r['담당자이메일']}<br>📞 {r['담당자연락처']}</td>
            <td><a href="{r['공고링크']}" target="_blank" style="background-color: #004aad; color: white; padding: 5px 10px; text-decoration: none; border-radius: 4px; font-size: 12px;">공고보기</a></td>
        </tr>
        """
    html_content += "</table></body></html>"
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[B2G 알림] {today_str} 전체 신규 지원사업 {len(results)}건 도착"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECEIVER_EMAILS)
    msg.attach(MIMEText(html_content, "html"))
    
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
    server.quit()
    print(f"🎉 전체 공고 {len(results)}건 이메일 단체 발송 성공!")

if __name__ == "__main__":
    notices = fetch_bizinfo_notices()
    analyzed_data = process_and_analyze(notices)
    send_email_report(analyzed_data)
