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
    # 💡 문제가 되었던 pageIndex를 완전히 제거하고 인증키와 검색 개수만 깔끔하게 요청합니다.
    params = {
        "crtfcKey": CRTFC_KEY,
        "dataType": "json",
        "searchCnt": "50"
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
    
    try:
        print("⏳ [진단 1단계] pageIndex를 제외하고 API 서버를 호출합니다...")
        session = requests.Session()
        resp = session.get(BIZINFO_API_URL, params=params, headers=headers, timeout=25, verify=False)
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
                print(f"❌ [데이터 형식 에러]: {json_e}")
                print(f"📦 [서버 원본 응답 내용 일부]: {resp.text[:300]}")
                return []
                
    except requests.exceptions.Timeout:
        print("❌ [에러 원인: 타임아웃] 서버 응답 지연")
    except requests.exceptions.ConnectionError:
        print("❌ [에러 원인: 연결 차단] 깃허브 IP 차단 또는 네트워크 불안정")
    except Exception as e:
        print(f"❌ [에러 발생]: {str(e)}")
        
    return []

def process_and_analyze(items):
    matched_results = []
    for item in items:
        title = item.get("pblancNm") or item.get("title", "")
        summary = item.get("bsnsSumryCn") or item.get("description", "")
        target_nm = item.get("trgetNm", "")
        org_name = item.get("jrsdInsttNm") or item.get("author", "-")
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
        print("\n⚠️ 수집된 공고가 없습니다.")
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
        print("❌ 공고를 가져오지 못해 분석을 종료합니다.")
