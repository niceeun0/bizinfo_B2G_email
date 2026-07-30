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
1. **OpenRouter 모델명**: `bizinfo_bot.py` 상단의 `OPENROUTER_MODEL` 기본값
   (`meta-llama/llama-3.1-8b-instruct:free`)은 예시입니다. openrouter.ai/models
   에서 현재 사용 가능한 무료 모델 슬러그를 확인 후, 워크플로우 yml의
   `OPENROUTER_MODEL` 환경변수 값을 교체하세요.
2. **기업마당 API 필드명**: 실제 API 응답 구조가 문서화되어 있지 않아, 코드
   내 `DATE_FIELD_CANDIDATES`, `ATTACHMENT_FIELD_CANDIDATES_*` 등은 여러
   후보 키를 자동으로 탐색하도록 방어적으로 작성했습니다. 실제 응답을 한 번
   출력해보고(`print(json.dumps(items[0], ensure_ascii=False, indent=2))`),
   필드명이 다르면 해당 리스트에 실제 키를 추가해주세요.
3. **네트워크 차단**: 워크플로우의 "DNS 상태 사전 점검" 스텝은 디버그용입니다.
   만약 curl 자체가 계속 실패한다면, 이는 로컬 DNS 문제가 아니라 서버가
   GitHub Actions의 IP 대역 자체를 차단하고 있다는 뜻이므로, 국내 리전
   프록시(예: 자체 VPS, 국내 클라우드) 경유가 필요합니다.

## 4. 로컬 테스트
```bash
pip install -r requirements.txt
export MAIL_USER=... EMAIL_PASS=... MAIL_RECEIVER=... GEMINI_API_KEY=...
python bizinfo_bot.py
```

## 5. 동작 요약
1. DoH로 IP를 직접 조회해 `curl --resolve`로 API 강제 호출 (실패 시 일반 curl → requests 순으로 폴백, 각 단계 지수 백오프 재시도)
2. 응답 JSON에서 **실행 전날** 등록된 공고만 필터링
3. 공고별 첨부파일 URL을 추출(HTML onclick 패턴 또는 경로+파일명 필드 조합)해 다운로드
4. Magic Bytes로 PDF/DOCX/HWP/HWPX 판별 후 각각 `pdfplumber` / `python-docx` / `olefile` / (스캔본은 `pytesseract` OCR)로 텍스트 추출
5. 정규식으로 담당자 이메일·전화번호 추출
6. OpenRouter로 지원규모/필수서류 2줄 요약
7. 카드형 HTML 뉴스레터 생성 (CSV Base64 다운로드 버튼 + mailto 원클릭 문의 버튼 포함)
8. Gmail SMTP(465, SSL)로 발송
