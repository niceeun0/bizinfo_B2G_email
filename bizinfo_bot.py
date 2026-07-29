# -*- coding: utf-8 -*-
import os
import re
import urllib.request
import json
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 💡 제공해주신 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

# 정규식 패턴 정의
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")
DOC_KEYWORDS = ["사업자등록증", "재무제표", "신용평가", "인감증명서", "법인등기", "주주명부", "국세완납", "지방세완납", "견적서", "소개서", "이력서", "신청서"]

def fetch_bizinfo_notices():
    # 💡 직접 API 호출 주소 생성
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=50"
    
    # SSL 인증 우회 설정 (서버 간 통신 안정성 확보)
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

    try:
        print("⏳ [수집 시작] 표준 보안 연결로 기업마당 API 호출 중...")
        # 30초 타임아웃 설정
        with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
            if response.status == 200:
                res_body = response.read().decode('utf-8')
                data = json.loads(res_body)
                items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                print(f"🎯 [수집 성공] 총 {len(items)}건의 공고를 가져왔습니다.")
                return items
            else:
                print(f"❌ [API 응답 오류]: 상태 코드 {response.status}")
    except Exception as e:
        print(f"❌ [연결 에러]: {str(e)}")
    return []

def analyze_and_build_html(items):
    rows_html = ""
    
    for idx, item in enumerate(items, 1):
        title = item.get("pblancNm") or item.get("title", "제목없음")
        exec_org = item.get("excInsttNm") or item.get("jrsdInsttNm") or "수행기관 미표기"
        period = item.get("reqstBeginEndDe") or item.get("reqstDt", "일정 참조")
        link = item.get("pblancUrl") or item.get("link", "#")
        summary = item.get("bsnsSumryCn", "")
        ref_name = item.get("refrncNm", "")
        original_method = item.get("reqstMthDscd") or item.get("reqstMthCn") or item.get("reqstMthPapersCn") or item.get("rcivMth") or "공고문 참조"
        
        full_text = f"{summary} {ref_name} {original_method}"

        # 1. 대상 분류 (기업 vs 개인)
        target_type = "👤 개인" if ("개인" in full_text and "기업" not in full_text) else "🏢 기업"

        # 2. 복수 이메일 전체 추출 (중복 제거)
        raw_emails = EMAIL_PATTERN.findall(full_text)
        emails = list(set(raw_emails))
        
        is_email_apply = "이메일" in original_method or "전자우편" in original_method or len(emails) > 0
        is_direct_apply = "방문" in original_method or "직접" in original_method or "우편" in original_method or "서면" in original_method

        # 3. 원클릭 제안 및 이메일 칸 구성
        one_click_html = "-"
        email_html = "-"

        if len(emails) > 0:
            email_html = "<br>".join([f"📧 {e}" for e in emails])
            one_click_html = "<br>".join([f'<a href="mailto:{e}" style="background:#e6f4ea; color:#137333; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:bold; text-decoration:none; display:inline-block; margin:2px 0;">⚡ 메일 제안하기</a>' for e in emails])
        elif is_direct_apply:
            one_click_html = '<span style="background:#fef7e0; color:#b06000; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:bold;">📁 방문/직접제출</span>'
        elif is_email_apply:
            one_click_html = '<span style="color:#137333; font-weight:bold; font-size:11px;">이메일 접수</span>'

        # 4. 필요 서류 및 문의처 파싱
        found_docs = [doc for doc in DOC_KEYWORDS if doc in full_text]
        docs_html = ", ".join([f'<span style="color:#c5221f; font-weight:bold;">{d}</span>' for d in found_docs]) if found_docs else "공고문 참조"
        
        phones = PHONE_PATTERN.findall(full_text)
        contact_html = f"📞 {phones[0]}" if phones else (ref_name if ref_name else "-")

        rows_html += f"""
        <tr>
            <td style="padding:10px; border-bottom:1px solid #eee; text-align:center; font-size:11px;">{target_type}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-weight:bold;">{title}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; color:#555;">{exec_org}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-size:12px;">{period}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-size:12px;">{original_method}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; text-align:center;">{one_click_html}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-size:12px;">{email_html}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-size:11px;">{docs_html}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; font-size:11px;">{contact_html}</td>
            <td style="padding:10px; border-bottom:1px solid #eee; text-align:center;"><a href="{link}" target="_blank" style="background:#004aad; color:white; padding:5px 8px; text-decoration:none; border-radius:4px; font-size:11px;">보기</a></td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family:'Malgun Gothic', sans-serif; color:#333;">
        <div style="max-width:1500px; margin:0 auto; padding:20px;">
            <h2 style="color:#004aad;">📊 B2G 지원사업 실시간 수집 및 원클릭 제안 리포트</h2>
            <p>기업마당 API 연동 결과, 총 {len(items)}건의 공고가 수집되었습니다.</p>
            <table style="width:100%; border-collapse:collapse; margin-top:15px;">
                <thead>
                    <tr style="background-color:#f8fafc; color:#444; font-size:12px;">
                        <th style="padding:10px; border-bottom:2px solid #ddd;">구분</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">공고명</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">사업수행기관</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">신청기간</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">신청방법 (원본 상세)</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">원클릭 제안</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">접수 이메일 (전체)</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">파싱된 필수 서류</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">담당 문의처</th>
                        <th style="padding:10px; border-bottom:2px solid #ddd;">바로가기</th>
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
    sender_password = os.environ.get("MAIL_PASS")
    receiver_list = os.environ.get("MAIL_RECEIVER") or sender_email

    if not sender_email or not sender_password:
        print("⚠️ [이메일 생략] 깃허브 Secrets에 MAIL_USER 또는 MAIL_PASS가 설정되지 않았습니다.")
        return

    receivers = [email.strip() for email in receiver_list.split(",")]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "📊 [자동화 리포트] B2G 지원사업 및 원클릭 제안 현황"
    msg["From"] = sender_email
    msg["To"] = ", ".join(receivers)
    
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receivers, msg.as_string())
        server.quit()
        print(f"🎉 [이메일 전송 성공] 다음 수신자에게 메일이 발송되었습니다: {receivers}")
    except Exception as e:
        print(f"❌ [이메일 전송 실패]: {str(e)}")

if __name__ == "__main__":
    notices = fetch_bizinfo_notices()
    if notices:
        html_report = analyze_and_build_html(notices)
        send_email(html_report)
    else:
        print("❌ 공고 데이터를 가져오지 못해 작업을 종료합니다.")
