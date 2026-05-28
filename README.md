# EJEP Camp 2026 — 체크인 시스템

## 파일 구조
```
camp-checkin/
├── server.js              # Express API 서버
├── package.json
├── Dockerfile
├── docker-compose.yml
├── data/
│   └── members.json       # 명단 데이터 (자동 생성/관리)
└── public/
    ├── index.html         # 기존 사이트 (체크인 섹션 교체됨)
    ├── checkin.html       # 셀프 체크인 페이지
    └── admin.html         # 관리자 페이지
```

---

## NAS 배포 방법 (Docker)

### 1. 파일 업로드
NAS의 원하는 폴더에 전체 파일을 복사합니다.

### 2. 환경변수 설정
`docker-compose.yml`에서 수정:
```yaml
environment:
  - ADMIN_PASSWORD=원하는비밀번호   # 관리자 비밀번호 변경!
  - PORT=3000
```

### 3. 실행
```bash
cd camp-checkin
docker-compose up -d
```

### 4. 접속 확인
- 메인 사이트: `http://NAS주소:3000`
- 셀프 체크인: `http://NAS주소:3000/checkin.html`
- 관리자:      `http://NAS주소:3000/admin.html`

---

## 엑셀 업로드 형식

관리자 페이지에서 엑셀 파일을 업로드할 수 있습니다.

| 컬럼명 | 필수 | 설명 |
|--------|------|------|
| 이름 | ✅ | 한글 이름 |
| 전화번호 | ✅ | 숫자만 (예: 0412345678) |
| 성별 | — | 남 / 여 |
| 나이 | — | 숫자 |
| 가족그룹ID | — | 같은 가족은 동일한 값 (예: FAM-001) |
| 방배정 | — | 방 번호/이름 |
| 팀배정 | — | 팀 이름 |
| 비고 | — | 식이제한, 특이사항 등 |

> 기존 체크인 기록이 있는 사람은 업데이트 시 체크인 정보가 유지됩니다.

---

## API 엔드포인트

| 메서드 | 경로 | 설명 |
|--------|------|------|
| POST | `/api/checkin` | 체크인 (이름 + 전화 끝 3자리) |
| GET | `/api/stats` | 총원/참석/결원 통계 (공개) |
| GET | `/api/admin/members` | 전체 명단 (관리자) |
| PATCH | `/api/admin/members/:id` | 체크인 상태 변경 (관리자) |
| POST | `/api/admin/upload` | 엑셀 업로드 (관리자) |
| GET | `/api/admin/export` | CSV 내보내기 (관리자) |

관리자 API는 요청 헤더에 `x-admin-password` 필요.

---

## 관리자 비밀번호

기본값: `ejep2026` → **반드시 변경하세요!**
`docker-compose.yml`의 `ADMIN_PASSWORD` 값을 수정 후 재시작:
```bash
docker-compose down && docker-compose up -d
```
