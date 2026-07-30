#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bizinfo_bot.py
=========================================================
기업마당(BizInfo) 신규 지원사업 공고 자동 수집 -> 첨부파일 원문 추출
-> AI 요약(OpenRouter) -> HTML 뉴스레터 생성 -> Gmail SMTP 발송

GitHub Actions(Ubuntu-latest) 환경에서 실행되도록 설계되었습니다.

[중요 참고사항]
1. "Temporary failure in name resolution" / "Connection timed out" 에러는
   대부분 GitHub Actions 러너의 일시적 로컬 DNS 조회 실패이며, 서버가
   실제로 클라우드 IP 대역 자체를 차단하는 경우도 있을 수 있습니다.
   완전한 우회를 보장할 수는 없으나, 아래 방식으로 성공률을 크게
   높였습니다:
     - Google/Cloudflare DoH(DNS-over-HTTPS)로 직접 IP를 조회
     - curl --resolve 로 로컬 DNS를 건너뛰고 강제로 해당 IP에 연결
     - curl 실패 시 requests 로 재시도 (다른 TLS/HTTP 스택 사용)
     - 지수 백오프 재시도(최대 5회)
   그래도 실패한다면 서버 측에서 클라우드 IP 대역 자체를 차단하고
   있는 것이므로, 국내 리전 프록시/자체 서버 경유가 필요합니다.

2. 기업마당 API의 실제 JSON 필드명은 시점에 따라 다를 수 있습니다.
   아래 코드는 흔히 쓰이는 필드명(pblancId, pblancNm, jrsdInsttNm,
   bsnsSumryCn, reqstBeginEndDe, creatPnttm, pubDate, printFlpthNm,
   printFileNm, rceptEngnHmpgUrl 등)을 우선 시도하고, 없으면 유사한
   키를 탐색하도록 방어적으로 작성했습니다. 실제 응답 구조가 다르면
   FIELD_CANDIDATES 부분만 수정하면 됩니다.

3. 환경변수 GEMINI_API_KEY 에는 (요청하신 대로 변수명은 유지하되)
   실제로는 OpenRouter API 키를 넣어서 사용합니다.
=========================================================
"""

import os
import re
import io
import csv
import sys
import json
import time
import base64
import socket
import zipfile
import smtplib
import subprocess
import traceback
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

# ---- 선택적 문서 파싱 라이브러리 (없으면 해당 포맷만 건너뜀) ----
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import docx  # python-docx
except ImportError:
    docx = None

try:
    import olefile
except ImportError:
    olefile = None

try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None


# =========================================================
# 0. 설정 / 환경변수
# =========================================================
BIZINFO_HOST = "www.bizinfo.go.kr"
BIZINFO_API_URL = (
    "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
    "?crtfcKey=4vc2gy&dataType=json&searchCnt=100"
)

MAIL_USER = os.environ.get("MAIL_USER", "")
EMAIL_PASS = os.environ.get("EMAIL_PASS", "")
MAIL_RECEIVER = os.environ.get("MAIL_RECEIVER", "")

# 요청사항: OpenRouter를 쓰되 환경변수 이름은 GEMINI_API_KEY 그대로 사용
OPENROUTER_API_KEY = os.environ.get("GEMINI_API_KEY", "")
# openrouter/free 는 OpenRouter가 공식 지원하는 "Free Models Router" 슬러그로,
# 요청 특성에 맞는 무료 모델을 자동으로 골라 라우팅합니다.
# (특정 모델을 고정하고 싶다면 "모델명:free" 형태로 환경변수 override 가능,
#  단 개별 무료 모델 슬러그는 자주 폐기/변경되므로 openrouter/free 권장)
OPENROUTER_MODEL = os.environ.get("OPENROUTER_MODEL", "openrouter/free")
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

REQUEST_TIMEOUT = 20
MAX_RETRY = 5
BACKOFF_BASE = 3  # seconds

WORKDIR = os.path.abspath(os.path.dirname(__file__))
TMP_DIR = os.path.join(WORKDIR, "_tmp_attachments")
os.makedirs(TMP_DIR, exist_ok=True)


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


# =========================================================
# 1. 네트워크 우회 / 견고한 API 호출
# =========================================================
def resolve_ip_via_doh(hostname):
    """Google / Cloudflare DoH를 이용해 로컬 DNS를 건너뛰고 IP를 직접 조회."""
    doh_endpoints = [
        f"https://dns.google/resolve?name={hostname}&type=A",
        f"https://cloudflare-dns.com/dns-query?name={hostname}&type=A",
    ]
    for url in doh_endpoints:
        try:
            headers = {"accept": "application/dns-json"}
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                answers = data.get("Answer", [])
                for ans in answers:
                    ip = ans.get("data")
                    # A 레코드(IPv4)만 사용
                    if ip and re.match(r"^\d{1,3}(\.\d{1,3}){3}$", ip):
                        log(f"DoH로 IP 확인 성공: {hostname} -> {ip} (via {url})")
                        return ip
        except Exception as e:
            log(f"DoH 조회 실패 ({url}): {e}")
            continue
    return None


def curl_get(url, resolved_ip=None, timeout=REQUEST_TIMEOUT):
    """subprocess로 curl 호출. resolved_ip가 있으면 --resolve로 강제 연결."""
    cmd = [
        "curl",
        "-sS",
        "-L",
        "--max-time", str(timeout),
        "--connect-timeout", str(min(timeout, 10)),
        "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) BizInfoBot/1.0",
    ]
    if resolved_ip:
        cmd += ["--resolve", f"{BIZINFO_HOST}:443:{resolved_ip}"]
    cmd += [url]

    result = subprocess.run(cmd, capture_output=True, timeout=timeout + 5)
    if result.returncode != 0:
        raise RuntimeError(
            f"curl 실패 (returncode={result.returncode}): "
            f"{result.stderr.decode('utf-8', errors='ignore')[:300]}"
        )
    return result.stdout


def fetch_with_retry(url):
    """
    다단계 견고한 호출:
      1) curl + DoH로 조회한 IP를 --resolve로 강제 사용
      2) curl (일반 시스템 DNS)
      3) requests (일반 시스템 DNS, 다른 네트워크 스택)
    각 단계 실패 시 지수 백오프로 재시도.
    """
    resolved_ip = resolve_ip_via_doh(BIZINFO_HOST)

    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        # 1) curl + --resolve (DoH IP)
        if resolved_ip:
            try:
                log(f"[시도 {attempt}] curl --resolve 방식으로 호출")
                raw = curl_get(url, resolved_ip=resolved_ip)
                return raw
            except Exception as e:
                last_err = e
                log(f"curl --resolve 실패: {e}")

        # 2) curl 일반 호출
        try:
            log(f"[시도 {attempt}] curl 일반 호출")
            raw = curl_get(url, resolved_ip=None)
            return raw
        except Exception as e:
            last_err = e
            log(f"curl 일반 호출 실패: {e}")

        # 3) requests 폴백
        try:
            log(f"[시도 {attempt}] requests 폴백 호출")
            resp = requests.get(
                url,
                timeout=REQUEST_TIMEOUT,
                headers={"User-Agent": "Mozilla/5.0 BizInfoBot/1.0"},
            )
            resp.raise_for_status()
            return resp.content
        except Exception as e:
            last_err = e
            log(f"requests 폴백 실패: {e}")

        sleep_time = BACKOFF_BASE * attempt
        log(f"{sleep_time}초 대기 후 재시도합니다...")
        time.sleep(sleep_time)

    raise RuntimeError(f"모든 재시도 실패. 마지막 에러: {last_err}")


def fetch_bizinfo_items():
    raw = fetch_with_retry(BIZINFO_API_URL)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"JSON 파싱 실패: {e} / 원본 앞부분: {raw[:200]}")

    # 응답 구조 방어적 처리: jsonArray 혹은 최상위 리스트 등
    if isinstance(data, dict):
        items = data.get("jsonArray") or data.get("items") or data.get("result") or []
    elif isinstance(data, list):
        items = data
    else:
        items = []

    log(f"API에서 총 {len(items)}건의 공고를 수신했습니다.")
    return items


# =========================================================
# 2. 날짜 필터링 (메일 발송 전날 등록된 신규 공고만)
# =========================================================
DATE_FIELD_CANDIDATES = ["pubDate", "creatPnttm", "regDt", "creatPnttm1", "rceptDt"]


def get_item_date_str(item):
    for key in DATE_FIELD_CANDIDATES:
        if item.get(key):
            return str(item[key])
    return None


def parse_flexible_date(date_str):
    """다양한 형식(YYYYMMDD, YYYY-MM-DD, RFC1123 등)을 안전하게 파싱."""
    if not date_str:
        return None
    date_str = date_str.strip()

    # RFC 1123 (RSS pubDate 형식) 예: "Wed, 29 Jul 2026 10:00:00 +0900"
    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(date_str)
        if dt:
            return dt.date()
    except Exception:
        pass

    candidates = [
        ("%Y-%m-%d %H:%M:%S", 19),
        ("%Y-%m-%d", 10),
        ("%Y.%m.%d", 10),
        ("%Y/%m/%d", 10),
        ("%Y%m%d", 8),
    ]
    for fmt, length in candidates:
        try:
            return datetime.strptime(date_str[:length], fmt).date()
        except Exception:
            continue

    # 마지막 시도: 숫자만 추출해서 YYYYMMDD로 처리
    digits = re.sub(r"\D", "", date_str)
    if len(digits) >= 8:
        try:
            return datetime.strptime(digits[:8], "%Y%m%d").date()
        except Exception:
            pass
    return None


def filter_new_items(items, target_date):
    """target_date(등록일)에 등록된 공고만 필터링."""
    filtered = []
    for item in items:
        date_str = get_item_date_str(item)
        item_date = parse_flexible_date(date_str)
        if item_date == target_date:
            filtered.append(item)
    log(f"{target_date} 등록 공고 {len(filtered)}건 필터링 완료.")
    return filtered


# =========================================================
# 3. 첨부파일 URL 추출
# =========================================================
FILEBLANK_RE = re.compile(
    r"fileBlank\(\s*'([^']+)'\s*\+\s*'([^']*)'\s*\+\s*'([^']+)'\s*,\s*'([^']+)'\s*\)"
)
FILEBLANK_SIMPLE_RE = re.compile(
    r"fileBlank\(\s*'([^']+)'\s*,\s*'([^']+)'\s*\)"
)

ATTACHMENT_FIELD_CANDIDATES_PATH = ["printFlpthNm", "flpthNm", "atchFileUrl"]
ATTACHMENT_FIELD_CANDIDATES_NAME = ["printFileNm", "fileNm", "atchFileNm"]
ATTACHMENT_RAW_HTML_CANDIDATES = ["fileDvNm", "atchFileHtml", "flieDown"]


def extract_attachment_url(item):
    """
    첨부파일 다운로드 URL과 원본 파일명을 반환. 없으면 (None, None).

    우선순위:
      1) onclick="fileBlank('path1' + 'path2' + 'path3', 'display_name.ext')"
         형태의 HTML이 필드 값에 그대로 들어있는 경우 정규식으로 파싱
      2) printFlpthNm(경로) + printFileNm(파일명) 필드 조합
    """
    # 1) HTML onclick 패턴이 필드 어딘가에 통째로 들어있는 경우
    for key in ATTACHMENT_RAW_HTML_CANDIDATES + ATTACHMENT_FIELD_CANDIDATES_PATH:
        val = item.get(key)
        if not val or not isinstance(val, str):
            continue
        if "fileBlank(" not in val:
            continue

        m = FILEBLANK_RE.search(val)
        if m:
            path = (m.group(1) + m.group(2) + m.group(3)).replace("//", "/")
            if not path.startswith("/"):
                path = "/" + path
            display_name = m.group(4)
            return f"https://{BIZINFO_HOST}{path}", display_name

        m2 = FILEBLANK_SIMPLE_RE.search(val)
        if m2:
            path = m2.group(1)
            if not path.startswith("/"):
                path = "/" + path
            display_name = m2.group(2)
            return f"https://{BIZINFO_HOST}{path}", display_name

    # 2) 경로 + 파일명 필드 조합 (일반적인 기업마당 API 구조)
    path_val = None
    for key in ATTACHMENT_FIELD_CANDIDATES_PATH:
        if item.get(key):
            path_val = str(item[key])
            break

    name_val = None
    for key in ATTACHMENT_FIELD_CANDIDATES_NAME:
        if item.get(key):
            name_val = str(item[key])
            break

    if path_val:
        # 이미 완전한 URL인 경우
        if path_val.startswith("http"):
            return path_val, (name_val or os.path.basename(path_val))
        # 상대 경로인 경우 도메인 결합, 중복 슬래시 정리
        full = f"https://{BIZINFO_HOST}/" + path_val.lstrip("/")
        full = re.sub(r"(?<!:)//+", "/", full.replace("https:/", "https://"))
        return full, (name_val or os.path.basename(path_val))

    return None, None


def download_attachment(url):
    """curl 우선, 실패 시 requests로 첨부파일 바이트 다운로드."""
    try:
        cmd = [
            "curl", "-sSL", "--max-time", "30",
            "-A", "Mozilla/5.0 BizInfoBot/1.0",
            url,
        ]
        result = subprocess.run(cmd, capture_output=True, timeout=35)
        if result.returncode == 0 and result.stdout:
            return result.stdout
    except Exception as e:
        log(f"첨부파일 curl 다운로드 실패: {e}")

    try:
        resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        log(f"첨부파일 requests 다운로드 실패: {e}")
        return None


# =========================================================
# 4. Magic Bytes 판독 + 텍스트 추출
# =========================================================
def detect_file_type(raw_bytes):
    if not raw_bytes or len(raw_bytes) < 8:
        return "unknown"
    header = raw_bytes[:8]

    if header[:4] == b"%PDF":
        return "pdf"

    if header[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
        return "hwp"  # 구버전 HWP (OLE Compound File)

    if header[:4] == b"PK\x03\x04":
        # zip 컨테이너 -> docx / hwpx 구분 필요
        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
                names = zf.namelist()
                if any(n.startswith("word/") for n in names):
                    return "docx"
                if any("Contents/section" in n for n in names) or "mimetype" in names:
                    return "hwpx"
        except Exception:
            pass
        return "zip_unknown"

    return "unknown"


def extract_text_from_pdf(raw_bytes):
    if not pdfplumber:
        return ""
    text_parts = []
    tmp_path = os.path.join(TMP_DIR, f"tmp_{int(time.time()*1000)}.pdf")
    with open(tmp_path, "wb") as f:
        f.write(raw_bytes)
    try:
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                text_parts.append(t)
        text = "\n".join(text_parts).strip()

        # 텍스트가 거의 없으면(스캔본 PDF) OCR 시도
        if len(text) < 20 and pytesseract and convert_from_path:
            log("PDF 텍스트가 거의 없어 OCR을 시도합니다.")
            try:
                images = convert_from_path(tmp_path, dpi=200)
                ocr_texts = []
                for img in images[:5]:  # 과도한 처리 방지: 최대 5페이지
                    ocr_texts.append(pytesseract.image_to_string(img, lang="kor+eng"))
                text = "\n".join(ocr_texts).strip()
            except Exception as e:
                log(f"OCR 실패: {e}")
        return text
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def extract_text_from_docx(raw_bytes):
    if not docx:
        return ""
    try:
        f = io.BytesIO(raw_bytes)
        document = docx.Document(f)
        paragraphs = [p.text for p in document.paragraphs]
        # 표 안의 텍스트도 함께 추출
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    paragraphs.append(cell.text)
        return "\n".join(paragraphs).strip()
    except Exception as e:
        log(f"DOCX 추출 실패: {e}")
        return ""


def extract_text_from_hwp(raw_bytes):
    """
    구버전 HWP(OLE) 파일에서 PrvText 스트림을 읽어 텍스트를 추출합니다.
    PrvText는 '미리보기 텍스트'이므로 전체 본문이 아닌 요약 수준일 수
    있습니다. 완전한 본문 추출이 필요하면 pyhwp(hwp5) 라이브러리 사용을
    권장합니다.
    """
    if not olefile:
        return ""
    try:
        f = io.BytesIO(raw_bytes)
        if not olefile.isOleFile(f):
            return ""
        ole = olefile.OleFileIO(f)
        if ole.exists("PrvText"):
            data = ole.openstream("PrvText").read()
            text = data.decode("utf-16-le", errors="ignore")
            return text.strip()
        return ""
    except Exception as e:
        log(f"HWP 추출 실패: {e}")
        return ""


def extract_text_from_hwpx(raw_bytes):
    """HWPX는 zip 기반 XML 포맷. section*.xml 내 텍스트 노드를 추출합니다."""
    try:
        text_parts = []
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            section_files = sorted(
                [n for n in zf.namelist() if re.search(r"Contents/section\d*\.xml$", n)]
            )
            for name in section_files:
                xml_content = zf.read(name).decode("utf-8", errors="ignore")
                # <hp:t>...</hp:t> 또는 <t>...</t> 텍스트 노드 추출
                chunks = re.findall(r"<(?:hp:)?t[^>]*>(.*?)</(?:hp:)?t>", xml_content, re.DOTALL)
                cleaned = [re.sub(r"<[^>]+>", "", c) for c in chunks]
                text_parts.extend(cleaned)
        return "\n".join(text_parts).strip()
    except Exception as e:
        log(f"HWPX 추출 실패: {e}")
        return ""


def extract_text_from_attachment(raw_bytes):
    if not raw_bytes:
        return "", "unknown"
    file_type = detect_file_type(raw_bytes)
    log(f"첨부파일 타입 감지 결과: {file_type}")

    if file_type == "pdf":
        return extract_text_from_pdf(raw_bytes), file_type
    if file_type == "docx":
        return extract_text_from_docx(raw_bytes), file_type
    if file_type == "hwp":
        return extract_text_from_hwp(raw_bytes), file_type
    if file_type == "hwpx":
        return extract_text_from_hwpx(raw_bytes), file_type
    return "", file_type


# =========================================================
# 5. 이메일 / 전화번호 정규식 추출
# =========================================================
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4})")


def extract_contacts(text):
    if not text:
        return [], []
    emails = sorted(set(EMAIL_RE.findall(text)))
    phones = sorted(set(PHONE_RE.findall(text)))
    return emails, phones


# =========================================================
# 6. OpenRouter AI 요약
# =========================================================
def summarize_with_openrouter(title, body_text):
    """
    원문에서 [지원규모]와 [필수서류]를 한국어 두 줄로 요약.
    API 키가 없거나 호출 실패 시, 안전한 대체 문구를 반환합니다.
    """
    fallback = (
        "📌 지원규모: 원문 확인이 필요합니다 (본문 추출/요약 실패).\n"
        "📋 필수서류: 공고 원문의 신청 서류 안내를 확인해 주세요."
    )

    if not OPENROUTER_API_KEY:
        log("OPENROUTER_API_KEY(GEMINI_API_KEY) 미설정 - 요약을 건너뜁니다.")
        return fallback

    if not body_text or len(body_text.strip()) < 10:
        return fallback

    prompt = (
        "다음은 정부/공공기관 지원사업 공고 원문입니다. "
        "이 내용을 바탕으로 정확히 아래 두 줄 형식으로만 한국어로 요약해줘. "
        "형식 이외의 다른 설명은 절대 추가하지 마.\n\n"
        "📌 지원규모: (한 문장)\n"
        "📋 필수서류: (한 문장, 쉼표로 나열 가능)\n\n"
        f"[공고 제목]\n{title}\n\n"
        f"[공고 원문 일부]\n{body_text[:3000]}"
    )

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 300,
        "temperature": 0.3,
    }

    for attempt in range(1, 4):
        try:
            resp = requests.post(
                OPENROUTER_URL, headers=headers, json=payload, timeout=30
            )
            if resp.status_code != 200:
                # 원인 진단을 위해 상태코드와 응답 본문 일부를 로그에 남김
                log(
                    f"OpenRouter 요약 실패(시도 {attempt}): "
                    f"HTTP {resp.status_code} - {resp.text[:300]}"
                )
                resp.raise_for_status()
            data = resp.json()
            content = data["choices"][0]["message"]["content"].strip()
            if content:
                return content
        except Exception as e:
            log(f"OpenRouter 요약 실패(시도 {attempt}): {e}")
            time.sleep(2 * attempt)

    return fallback


# =========================================================
# 7. HTML 뉴스레터 생성
# =========================================================
def build_csv_base64(items_data):
    """전체 수집 데이터를 CSV로 만들고 Base64 인코딩."""
    output = io.StringIO()
    fieldnames = [
        "제목", "소관기관", "등록일", "신청기간", "요약",
        "담당 이메일", "담당 전화번호", "첨부파일명", "원문링크",
    ]
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    for row in items_data:
        writer.writerow(row)

    csv_bytes = output.getvalue().encode("utf-8-sig")  # 엑셀 한글 깨짐 방지 BOM
    b64 = base64.b64encode(csv_bytes).decode("ascii")
    return b64


def build_html_newsletter(items_data, target_date):
    csv_b64 = build_csv_base64(items_data)
    filename = f"bizinfo_{target_date.strftime('%Y%m%d')}.csv"

    cards_html = ""
    for row in items_data:
        title = row["제목"]
        agency = row["소관기관"]
        period = row["신청기간"]
        summary_html = row["요약"].replace("\n", "<br>")
        email = row["담당 이메일"]
        phone = row["담당 전화번호"]
        link = row["원문링크"] or "#"

        mailto_btn = ""
        if email:
            subject = requests.utils.quote(f"[지원사업 문의] {title}")
            body = requests.utils.quote(
                f"안녕하세요, '{title}' 공고 관련 문의드립니다.\n\n"
                f"- 문의 내용:\n"
            )
            mailto_btn = (
                f'<a href="mailto:{email}?subject={subject}&body={body}" '
                f'style="display:inline-block;margin-top:8px;padding:6px 14px;'
                f'background:#2563eb;color:#fff;border-radius:6px;'
                f'text-decoration:none;font-size:13px;">✉️ 원클릭 메일 문의</a>'
            )

        cards_html += f"""
        <div style="border:1px solid #e5e7eb;border-radius:10px;padding:16px 18px;
                    margin-bottom:14px;background:#ffffff;
                    box-shadow:0 1px 3px rgba(0,0,0,0.05);">
          <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:4px;">
            {title}
          </div>
          <div style="font-size:12.5px;color:#6b7280;margin-bottom:10px;">
            🏢 {agency} &nbsp;|&nbsp; 📅 신청기간: {period}
          </div>
          <div style="font-size:13.5px;color:#111827;line-height:1.6;
                      background:#f9fafb;border-radius:6px;padding:10px 12px;
                      margin-bottom:8px;">
            {summary_html}
          </div>
          <div style="font-size:12.5px;color:#374151;">
            {"📧 " + email if email else ""} {" &nbsp; " if email and phone else ""}
            {"📞 " + phone if phone else ""}
          </div>
          <div style="margin-top:6px;">
            <a href="{link}" style="font-size:12.5px;color:#2563eb;text-decoration:none;
                      margin-right:10px;">🔗 원문 바로가기</a>
            {mailto_btn}
          </div>
        </div>
        """

    html = f"""
    <html>
    <head><meta charset="utf-8"></head>
    <body style="margin:0;padding:0;background:#f3f4f6;font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif;">
      <div style="max-width:640px;margin:0 auto;padding:24px 16px;">
        <div style="text-align:center;margin-bottom:20px;">
          <h2 style="color:#111827;margin:0 0 4px 0;">📢 기업마당 신규 지원사업 뉴스레터</h2>
          <div style="color:#6b7280;font-size:13px;">
            {target_date.strftime('%Y년 %m월 %d일')} 등록 공고 &middot; 총 {len(items_data)}건
          </div>
        </div>

        {cards_html if items_data else '<div style="text-align:center;color:#6b7280;padding:40px 0;">오늘 등록된 신규 공고가 없습니다.</div>'}

        <div style="text-align:center;margin-top:20px;">
          <a download="{filename}"
             href="data:text/csv;base64,{csv_b64}"
             style="display:inline-block;padding:10px 20px;background:#111827;
                    color:#fff;border-radius:8px;text-decoration:none;
                    font-size:13.5px;">
            📥 전체 데이터 엑셀(CSV) 다운로드
          </a>
        </div>

        <div style="text-align:center;color:#9ca3af;font-size:11px;margin-top:24px;">
          본 메일은 기업마당(bizinfo.go.kr) 공개 API를 기반으로 자동 생성되었습니다.
        </div>
      </div>
    </body>
    </html>
    """
    return html


# =========================================================
# 8. Gmail SMTP 발송
# =========================================================
def parse_recipients(raw_receiver):
    """
    MAIL_RECEIVER 값을 안전하게 파싱합니다.
    - 콤마(,) 또는 세미콜론(;)으로 구분된 다중 수신자 지원
    - "홍길동 <hong@test.com>" 같은 이름 포함 형식도 지원
      (SMTP envelope에는 순수 이메일 주소만, 헤더 표시는 이름 포함 유지)
    - RFC 5321에 맞지 않는(순수 이메일 형식이 아닌) 항목은 걸러내고 경고 로그 출력
    반환값: (envelope_emails: list[str], header_value: str)
    """
    from email.utils import getaddresses, formataddr

    if not raw_receiver:
        return [], ""

    # 세미콜론을 콤마로 통일한 뒤, 빈 토큰(끝/중간의 연속 콤마 등)을 먼저
    # 제거합니다. getaddresses는 빈 토큰이 섞이면 전체 파싱이 깨질 수 있습니다.
    normalized = raw_receiver.replace(";", ",")
    tokens = [t.strip() for t in normalized.split(",") if t.strip()]
    parsed = getaddresses(tokens)

    valid = []
    for name, addr in parsed:
        addr = addr.strip()
        if EMAIL_RE.fullmatch(addr):
            valid.append((name.strip(), addr))
        else:
            log(f"경고: 유효하지 않은 수신자 형식이라 제외합니다 -> name='{name}' addr='{addr}'")

    envelope_emails = [addr for _, addr in valid]
    header_value = ", ".join(
        formataddr((name, addr)) if name else addr for name, addr in valid
    )
    return envelope_emails, header_value


def send_email(subject, html_body):
    if not (MAIL_USER and EMAIL_PASS and MAIL_RECEIVER):
        raise RuntimeError(
            "MAIL_USER / EMAIL_PASS / MAIL_RECEIVER 환경변수가 모두 필요합니다."
        )

    envelope_emails, header_value = parse_recipients(MAIL_RECEIVER)
    if not envelope_emails:
        raise RuntimeError(
            "MAIL_RECEIVER에서 유효한 이메일 주소를 하나도 찾지 못했습니다. "
            f"원본 값(참고용): {MAIL_RECEIVER!r}"
        )
    log(f"발송 대상 수신자 {len(envelope_emails)}명 확인: {envelope_emails}")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = MAIL_USER
    msg["To"] = header_value
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    last_err = None
    for attempt in range(1, 4):
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as server:
                server.login(MAIL_USER, EMAIL_PASS)
                # envelope 수신자는 반드시 순수 이메일 주소 리스트여야 함
                server.sendmail(MAIL_USER, envelope_emails, msg.as_string())
            log("메일 발송 성공.")
            return
        except Exception as e:
            last_err = e
            log(f"메일 발송 실패(시도 {attempt}): {e}")
            time.sleep(3 * attempt)

    raise RuntimeError(f"메일 발송 최종 실패: {last_err}")


# =========================================================
# 9. 메인 파이프라인
# =========================================================
def main():
    log("=== 기업마당 뉴스레터 봇 시작 ===")

    # 실행 시각(KST) 기준, "전날" 등록 공고를 대상으로 함
    kst = timezone(timedelta(hours=9))
    now_kst = datetime.now(kst)
    target_date = (now_kst - timedelta(days=1)).date()
    log(f"대상 등록일(전날): {target_date}")

    try:
        raw_items = fetch_bizinfo_items()
    except Exception as e:
        log(f"[치명적 오류] API 호출 실패: {e}")
        traceback.print_exc()
        sys.exit(1)

    new_items = filter_new_items(raw_items, target_date)

    items_data = []
    for idx, item in enumerate(new_items, start=1):
        title = item.get("pblancNm") or item.get("bsnsNm") or "(제목 없음)"
        agency = item.get("jrsdInsttNm") or item.get("excInsttNm") or "-"
        period = item.get("reqstBeginEndDe") or item.get("pblancEndDe") or "-"
        link = item.get("pblancUrl") or item.get("rceptEngnHmpgUrl") or ""

        log(f"[{idx}/{len(new_items)}] 처리 중: {title}")

        attachment_url, attachment_name = extract_attachment_url(item)
        body_text = ""
        if attachment_url:
            log(f"  첨부파일 발견: {attachment_name} ({attachment_url})")
            raw_bytes = download_attachment(attachment_url)
            body_text, file_type = extract_text_from_attachment(raw_bytes)
            if not body_text:
                log("  첨부파일 텍스트 추출 실패 또는 빈 결과.")
        else:
            # 첨부파일이 없으면 공고 요약 필드라도 사용
            body_text = item.get("bsnsSumryCn") or ""

        emails, phones = extract_contacts(body_text)
        email = emails[0] if emails else ""
        phone = phones[0] if phones else ""

        summary = summarize_with_openrouter(title, body_text)

        items_data.append({
            "제목": title,
            "소관기관": agency,
            "등록일": str(target_date),
            "신청기간": period,
            "요약": summary,
            "담당 이메일": email,
            "담당 전화번호": phone,
            "첨부파일명": attachment_name or "",
            "원문링크": link,
        })

    html_body = build_html_newsletter(items_data, target_date)
    subject = f"[기업마당 신규공고] {target_date.strftime('%Y-%m-%d')} 등록 {len(items_data)}건"

    try:
        send_email(subject, html_body)
    except Exception as e:
        log(f"[치명적 오류] 메일 발송 실패: {e}")
        traceback.print_exc()
        sys.exit(1)

    log("=== 봇 실행 완료 ===")


if __name__ == "__main__":
    main()
