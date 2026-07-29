# -*- coding: utf-8 -*-
import io
import os
import re
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

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
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }
    
    try:
        print("⏳ [진단 1단계] 기업마당 API 서버로 접속을 시도합니다...")
        session = requests.Session()
        resp = session.get(req_url, headers=headers, timeout=25, verify=False)
        print(f"✅ [진단 통과] 서버 응답 상태 코드: {resp.status_code}")
        
        if resp.status_code == 200:
            try:
                data = resp.json()
                if "reqErr" in data:
                    print(f"❌ [API 서버 에러 발생]: {data['reqErr']}")
                    return []
                    
                items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                print(f"🎯 [파싱 성공] 총 {len(items)}건의 공고를 불러왔습니다.")
                return items
            except Exception as json_e:
                print(f"❌ [데이터 형식 에러]: 서버가 JSON이 아닌 다른 형태의 응답을 보냈습니다. 에러 내용: {json_e}")
                print(f"📦 [서버 원본 응답 내용 일부]: {resp.text[:300]}")
                return []
                
    except requests.exceptions.Timeout:
        print("❌ [에러 원인: 타임아웃] 기업마당 서버 응답 시간이 너무 오래 걸려 연결이 끊겼습니다.")
    except requests.exceptions.ConnectionError:
        print("❌ [에러 원인: 연결 거부/차단] 깃허브 IP가 기업마당 방화벽에 의해 차단되었거나 네트워크가 불안정합니다.")
    except Exception as e:
        print(f"❌ [알 수 없는 에러 발생]: {str(e)}")
        
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
        contact_info = item.get("refrncNm", "")
        papers_info = item.get("reqstMthPapersCn", "")
            
        target_text = f"{target_nm} {summary}"
        corp_score = sum(target_text.count(kw) for kw in CORP_KEYWORDS)
        ind_score = sum(target_text.count(kw) for kw in INDIVIDUAL_KEYWORDS)
        target_type = "🏢 기업 대상" if corp_score >= ind_score else "👤 개인 대상"
        
        extracted_docs = [doc for doc in DOC_KEYWORDS if doc in f"{summary} {papers_info}"]
        extracted_emails = set(EMAIL_PATTERN.findall(f"{summary} {contact_info}"))
        extracted_phones = set(PHONE_PATTERN.findall(f"{summary} {contact_info}"))
                
        matched_results.append({
            "공고명": title, "소관기관": org_name, 
            "지원대상": target_type, "신청기간": item.get("reqstBeginEndDe") or item.get("reqstDt", "-"),
            "필요서류": ", ".join(extracted_docs) if extracted_docs else "공고문 참조",
            "담당자연락처": f"📧 {', '.join(sorted(extracted_emails)) or '-'} / 📞 {', '.join(sorted(extracted_phones)) or '-'}",
            "공고링크": link
        })
    return matched_results

def print_screen_report(results):
    if not results:
        print("\n⚠️ 수집된 공고가 없어 화면 출력을 건너뜁니다.")
        return

    print("\n" + "="*80)
    print(" 📡 [기업마당] 실시간 수집 리포트 화면 출력 결과 ")
    print("="*80)
    
    for i, r in enumerate(results, 1):
        print(f"[{i}] {r['공고명']}")
        print(f" ├─ 구분: {r['지원대상']} | 소관기관: {r['소관기관']}")
        print(f" ├─ 신청기간: {r['신청기간']}")
        print(f" ├─ 필요서류: {r['필요서류']}")
        print(f" ├─ 담당자정보: {r['담당자연락처']}")
        print(f" └─ 링크: {r['공고링크']}")
        print("-" * 80)
    print("🎉 모든 공고 데이터 출력 완료!")

if __name__ == "__main__":
    notices = fetch_bizinfo_notices()
    if notices:
        analyzed_data = process_and_analyze(notices)
        print_screen_report(analyzed_data)
    else:
        print("❌ 공고를 가져오지 못해 분석을 종료합니다. 위 로그의 에러 메시지를 확인해 주세요.")
