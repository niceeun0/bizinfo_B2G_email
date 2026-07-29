# -*- coding: utf-8 -*-
"""
=======================================================================================
기업마당(Bizinfo) API 연동 B2G/정부지원사업 데일리 알림 봇 (bizinfo_bot.py)
=======================================================================================
"""

import io
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# 1. 설정 정보 (API 인증키 및 이메일 설정)
# ---------------------------------------------------------------------------
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "Ofgt6R"  # 대표님의 기업마당 인증키

# 사내 이메일 발송 세팅 (구글 웍스/아웃룩 등)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "your_email@gmail.com"  # 발송용 이메일
SENDER_PASSWORD = "your_app_password"  # 메일 앱 비밀번호
RECEIVER_EMAIL = "sales_team@ourdomain.com"  # 수신할 사내 영업팀 이메일

# 🎯 타겟팅 필터링 키워드
TARGET_KEYWORDS = [
    "데이터바우처", "마이데이터", "블록체인", "실증", 
    "기업정보", "기업DB", "데이터", "기업개요", "기업데이터"
]

# 🏢 기업 vs 👤 개인 대상 판별 키워드
CORP_KEYWORDS = ["중소기업", "기업", "법인", "사업자", "컨소시엄", "주관기관", "벤처", "소상공인"]
INDIVIDUAL_KEYWORDS = ["개인", "일반국민", "청년", "구직자", "학생", "개인사업자 제외"]

# 📄 제출서류 탐색 정규식 키워드
DOC_KEYWORDS = ["사업자등록증", "재무제표", "신용평가", "인감증명서", "법인등기", "주주명부", "국세완납", "지방세완납"]

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")

# ---------------------------------------------------------------------------
# 2. 기업마당 API 호출 및 1차 필터링
# ---------------------------------------------------------------------------
def fetch_bizinfo_notices():
    """기업마당 API를 호출하여 최신 지원사업 가져오기"""
    params = {
        "crtfcKey": CRTFC_KEY,
        "dataType": "json",
        "searchCnt": "100",  # 최근 100건 조회
        "pageIndex": "1"
    }
    
    try:
        resp = requests.get(BIZINFO_API_URL, params=params, timeout=15, verify=False)
        if resp.status_code != 200:
            print(f"API 호출 실패: {resp.status_code}")
            return []
        
        data = resp.json()
        items = data.get("jsonArray", []) or data.get("item", [])
        return items
    except Exception as e:
        print(f"API 요청 중 에러 발생: {e}")
        return []

# ---------------------------------------------------------------------------
# 3. 스마트 스캐닝 & 딥 파싱 엔진
# ---------------------------------------------------------------------------
def process_and_analyze(items):
    matched_results = []
    
    for item in items:
        # API 응답에서 핵심 항목 추출
        title = item.get("pblancNm") or item.get("title", "")
        summary = item.get("bsnsSumryCn") or item.get("description", "")
        target_nm = item.get("trgetNm", "")
        org_name = item.get("jrsdInsttNm") or item.get("author", "-")
        exec_org = item.get("excInsttNm", "-")
        link = item.get("pblancUrl") or item.get("link", "#")
        file_url = item.get("flpthNm") or item.get("printFlpthNm", "")
        contact_info = item.get("refrncNm", "")
        papers_info = item.get("reqstMthPapersCn", "")
        
        # 💡 [1단계 필터링] 공고명과 요약글에서 우리 키워드가 있는지 검사 (속도 최적화)
        full_text_for_kw = f"{title} {summary}"
        if not any(kw in full_text_for_kw for kw in TARGET_KEYWORDS):
            continue  # 상관없는 공고는 0.001초 만에 건너뜀 (서버 부하 차단)
            
        print(f"🔍 [타겟 공고 포착] {title}")
        
        # 💡 [2단계 분석] 기업 vs 개인 대상 분류
        target_text = f"{target_nm} {summary}"
        corp_score = sum(target_text.count(kw) for kw in CORP_KEYWORDS)
        ind_score = sum(target_text.count(kw) for kw in INDIVIDUAL_KEYWORDS)
        target_type = "🏢 기업 대상" if corp_score >= ind_score else "👤 개인 대상"
        
        # 💡 [3단계 분석] 제출 서류 및 담당자 정보 추출
        # API 텍스트에서 먼저 1차로 긁고, 첨부파일이 있으면 핀포인트 파싱
        extracted_docs = [doc for doc in DOC_KEYWORDS if doc in f"{summary} {papers_info}"]
        extracted_emails = set(EMAIL_PATTERN.findall(f"{summary} {contact_info}"))
        extracted_phones = set(PHONE_PATTERN.findall(f"{summary} {contact_info}"))
        
        # 첨부파일(PDF) 핀포인트 스캔 (필요 시 상위 2페이지만 가볍게 읽기)
        if file_url and file_url.endswith(".pdf"):
            try:
                f_resp = requests.get(file_url, timeout=10, verify=False)
                if f_resp.status_code == 200 and f_resp.content.startswith(b'%PDF'):
                    with pdfplumber.open(io.BytesIO(f_resp.content)) as pdf:
                        # 상위 2페이지만 스캔
                        pdf_text = "\n".join([p.extract_text() or "" for p in pdf.pages[:2]])
                        for doc in DOC_KEYWORDS:
                            if doc in pdf_text and doc not in extracted_docs:
                                extracted_docs.append(doc)
                        extracted_emails.update(EMAIL_PATTERN.findall(pdf_text))
                        extracted_phones.update(PHONE_PATTERN.findall(pdf_text))
            except Exception:
                pass  # 첨부파일 오류가 나도 시스템이 멈추지 않음
                
        matched_results.append({
            "공고명": title,
            "소관기관": org_name,
            "수행기관": exec_org,
            "지원대상": target_type,
            "신청기간": item.get("reqstBeginEndDe") or item.get("reqstDt", "-"),
            "필요서류": ", ".join(extracted_docs) if extracted_docs else "공고문 참조",
            "담당자이메일": ", ".join(sorted(extracted_emails)) or "-",
            "담당자연락처": ", ".join(sorted(extracted_phones)) or "-",
            "공고링크": link
        })
        
    return matched_results

# ---------------------------------------------------------------------------
# 4. 사내 이메일 HTML 발송 로직
# ---------------------------------------------------------------------------
def send_email_report(results):
    if not results:
        print(" 오늘 조건에 맞는 신규 지원사업이 없습니다.")
        return

    today_str = datetime.now().strftime("%Y년 %m월 %d일")
    
    html_content = f"""
    <html>
    <body style="font-family: 'Malgun Gothic', sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #004aad;">📡 [기업마당] 오늘의 B2G 데이터/지원사업 포착 리포트 ({today_str})</h2>
        <p>설정하신 <b>데이터, 기업개요, 기업데이터, 마이데이터, 블록체인</b> 등 핵심 키워드 공고입니다.</p>
        <hr style="border: 1px solid #ddd;">
        
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; text-align: left; font-size: 14px;">
            <tr style="background-color: #f2f4f7;">
                <th style="width: 12%;">구분</th>
                <th style="width: 28%;">공고명 / 소관기관</th>
                <th style="width: 15%;">신청기간</th>
                <th style="width: 20%;">필요 제출 서류</th>
                <th style="width: 15%;">담당자 연락처</th>
                <th style="width: 10%;">바로가기</th>
            </tr>
    """
    
    for r in results:
        html_content += f"""
        <tr>
            <td><b>{r['지원대상']}</b></td>
            <td>
                <b>{r['공고명']}</b><br>
                <span style="font-size: 12px; color: #666;">기관: {r['소관기관']} ({r['수행기관']})</span>
            </td>
            <td>{r['신청기간']}</td>
            <td><span style="color: #d93025; font-weight: bold;">{r['필요서류']}</span></td>
            <td>
                📧 {r['담당자이메일']}<br>
                📞 {r['담당자연락처']}
            </td>
            <td><a href="{r['공고링크']}" target="_blank" style="background-color: #004aad; color: white; padding: 5px 10px; text-decoration: none; border-radius: 4px; font-size: 12px;">공고보기</a></td>
        </tr>
        """
        
    html_content += """
        </table>
        <br>
        <p style="font-size: 12px; color: #888;">※ 본 메일은 B2G 데이터 영업 자동화 봇에 의해 생성되었습니다.</p>
    </body>
    </html>
    """
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[B2G 알림] {today_str} 신규 데이터/마이데이터/블록체인 지원사업 {len(results)}건 포착"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg.attach(MIMEText(html_content, "html"))
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
        server.quit()
        print("🎉 사내 이메일 발송 성공!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")

# ---------------------------------------------------------------------------
# 메인 실행부
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("🚀 기업마당 API 수집 및 스마트 분석 시작...")
    notices = fetch_bizinfo_notices()
    analyzed_data = process_and_analyze(notices)
    send_email_report(analyzed_data)