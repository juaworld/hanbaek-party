# 한백건축배 연말파티 접수 사이트

- `index.html` : 접수/수정/관리자 화면 (GitHub Pages로 배포)
- `Code.gs` : 구글시트에 붙이는 서버 코드 (Google Apps Script)
- 관리자 화면: 주소 끝에 `#admin`  (예: .../hanbaek-party/#admin)

## 설정
1. 구글시트 새로 만들기 → 확장 프로그램 → Apps Script → `Code.gs` 붙여넣기
2. `ADMIN_PASSWORD`, `SALT` 변경 후 저장
3. 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자) → URL 복사
4. `index.html` 맨 위 `const API_URL = '';` 에 URL 붙여넣고 커밋
