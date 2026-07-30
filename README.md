# 기업마당(BizInfo) 신규공고 뉴스레터 봇

## 1. 파일 구성
- `bizinfo_bot.py` : 메인 스크립트
- `requirements.txt` : 파이썬 의존성
- `.github/workflows/bizinfo_bot.yml` : GitHub Actions 워크플로우

## 2. GitHub Secrets 등록 (Settings → Secrets and variables → Actions)
| Secret 이름 | 설명 |
|---|---|
| `MAIL_USER` | 발신용 Gmail 주소 |
| `EMAIL_PASS` | Gmail **앱 비밀번호** (일반 로그인 비밀번호 아님, 2단계 인증 후 앱 비밀번호 발급 필요) |
| `MAIL_RECEIVER` | 수신 이메일 주소 (여러 명이면 콤마로 구분 후 코드에서 split 처리 필요) |
| `GEMINI_API_KEY` | **OpenRouter API 키**를 이 이름의 시크릿에 저장 (요청하신 변수명 유지) |

## 3. 반드시 확인/수정해야 할 부분
1. **OpenRouter 모델명**: 기본값은 `openrouter/free`입니다. 이는 OpenRouter가
   공식 지원하는 "Free Models Router" 슬러그로, 요청 특성에 맞는 무료 모델을
   자동으로 골라줍니다. 특정 모델을 고정하고 싶다면 openrouter.ai/models 에서
   현재 사용 가능한 `모델명:free` 슬러그를 확인 후 워크플로우 yml의
   `OPENROUTER_MODEL` 값을 교체하세요. 단, 개별 무료 모델 슬러그는 자주
   폐기/변경되어 404 에러가 나기 쉬우므로 `openrouter/free` 사용을 권장합니다.
2. **`MAIL_RECEIVER` 형식**: 순수 이메일 주소만 넣거나(`a@b.com`), 콤마 또는
   세미콜론으로 구분한 다중 수신자(`a@b.com,c@d.com`), `이름 <이메일>` 형식
   모두 지원합니다. 코드가 자동으로 파싱해서 유효한 주소만 SMTP 전송에
   사용하고, 유효하지 않은 항목은 로그에 경고로 남기고 건너뜁니다. 만약 모든
   항목이 걸러졌다면 실행 로그의 "유효하지 않은 수신자 형식이라 제외합니다"
   메시지로 원인을 확인하세요.
3. **기업마당 API 필드명**: 실제 API 응답 구조가 문서화되어 있지 않아, 코드
   내 `DATE_FIELD_CANDIDATES`, `ATTACHMENT_FIELD_CANDIDATES_*` 등은 여러
   후보 키를 자동으로 탐색하도록 방어적으로 작성했습니다. 실제 응답을 한 번
   출력해보고(`print(json.dumps(items[0], ensure_ascii=False, indent=2))`),
   필드명이 다르면 해당 리스트에 실제 키를 추가해주세요.
4. **네트워크 차단**: 워크플로우의 "DNS 상태 사전 점검" 스텝은 디버그용입니다.
   만약 curl 자체가 계속 실패한다면, 이는 로컬 DNS 문제가 아니라 서버가
   GitHub Actions의 IP 대역 자체를 차단하고 있다는 뜻이므로, 국내 리전
   프록시(예: 자체 VPS, 국내 클라우드) 경유가 필요합니다.

## 4. 로컬 테스트
```bash
pip install -r requirements.txt
export MAIL_USER=... EMAIL_PASS=... MAIL_RECEIVER=... GEMINI_API_KEY=...
python bizinfo_bot.py
```

## 5. 이번 업데이트 요약
1. **키워드/서류명 확장**: "최종선정/사업규모/모집인원" 등 선정 규모 표현과, 실제
   공고문에 자주 나오는 구체적 서류명(사업자등록증명, 중소기업확인서, 국세완납
   증명서 등)을 정규식/문자열로 직접 탐지합니다. AI 호출이 실패해도 이 자동
   감지 결과로 최소한의 요약을 제공하고, AI 호출 시에는 근거로 함께 제공해
   환각을 줄입니다.
2. **응답 정제**: 일부 무료 모델이 "USER SAFETY: SAFE" 같은 영어 메타 문구를
   응답에 섞어 보내는 문제를, 응답에서 📌/📋 두 줄만 정확히 추출해 사용하도록
   고쳐서 원천적으로 해결했습니다.
3. **유사 공고 자동 통합**: 제목 유사도(`DEDUP_TITLE_THRESHOLD`, 기본 0.55)와
   본문 유사도(`DEDUP_BODY_THRESHOLD`, 기본 0.92)가 모두 임계값을 넘을 때만
   중복으로 판단해 하나로 합칩니다. **다만 완전히 다른 세부 프로그램(예: 같은
   상위 사업의 "맞춤형/마케팅/디자인" 트랙)도 공통 보일러플레이트 문구가 많으면
   본문 유사도가 높게 나올 수 있어, 과도하게 합쳐지는 것 같으면 로그의
   "유사 공고로 판단되어 통합" 메시지에서 유사도 점수를 확인하고
   `DEDUP_BODY_THRESHOLD`를 0.95~0.97처럼 더 높여 보수적으로 조정하세요.**
4. **원클릭 서비스 제안 버튼**: 기존 "메일 문의" 버튼을 "📩 원클릭 서비스 제안"
   버튼으로 변경했습니다. 받는사람은 정규식으로 재검증된 순수 이메일 주소만
   들어가고, 본문은 서비스 제안 톤의 기본 문구가 들어갑니다(구체적인 제안
   내용은 `build_html_newsletter` 함수의 `body = urlquote(...)` 부분에서 직접
   다듬으시면 됩니다).
5. **실제 엑셀(.xlsx) 다운로드**: CSV 위장 다운로드 대신 `openpyxl`로 진짜
   `.xlsx` 파일을 만들어 제공합니다(헤더 스타일링, 열 너비, 틀 고정 포함).
   `openpyxl`이 설치되어 있지 않은 환경에서는 자동으로 CSV로 폴백합니다.

## 6. 동작 요약
1. DoH로 IP를 직접 조회해 `curl --resolve`로 API 강제 호출 (실패 시 일반 curl → requests 순으로 폴백, 각 단계 지수 백오프 재시도)
2. 응답 JSON에서 **실행 전날** 등록된 공고만 필터링
3. 공고별 첨부파일 URL을 추출(HTML onclick 패턴 또는 경로+파일명 필드 조합)해 다운로드
4. Magic Bytes로 PDF/DOCX/HWP/HWPX 판별 후 각각 `pdfplumber` / `python-docx` / `olefile` / (스캔본은 `pytesseract` OCR)로 텍스트 추출
5. 정규식으로 담당자 이메일·전화번호 추출
6. OpenRouter로 지원규모/필수서류 2줄 요약
7. 카드형 HTML 뉴스레터 생성 (CSV Base64 다운로드 버튼 + mailto 원클릭 문의 버튼 포함)
8. Gmail SMTP(465, SSL)로 발송
