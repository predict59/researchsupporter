# 국군복지단 가격조사 도우미

현장에서 조사표의 매장과 물품을 확인하며 가격, 판매 상태, 사진, 특이사항을 기록하는 React 기반 오프라인 우선 PWA입니다. 결과는 지역별 엑셀, 사진 ZIP, 백업 JSON으로 내보낼 수 있습니다.

## 주요 기능

- 조사표 엑셀과 선택형 업체 연락처 엑셀 불러오기
- 번들 기준자료의 상품사진 연결 및 바코드 이미지 자동 생성
- 지역/담당매장/물품별 조사 진행 관리
- 가격, 진열 여부, POS 확인, 사진 및 특이사항 입력
- 지도와 현재 위치를 이용한 매장 방문 지원
- 지역별 결과 엑셀과 사진 ZIP 내보내기
- 전체 작업 데이터 JSON 백업 및 복원

현장 데이터와 사진은 서버가 아니라 사용 중인 브라우저의 IndexedDB에 저장됩니다. 다른 기기나 브라우저로 자동 동기화되지 않으므로 전체 백업 JSON을 사용해야 합니다.

## 실행

```powershell
npm install
npm run dev
```

프로덕션 빌드 검증:

```powershell
npm run build
```

## GitHub 동기화와 배포

수정 사항을 빌드 검증한 뒤 `main`에 커밋하고 GitHub로 푸시하려면 다음 명령을 사용합니다.

```powershell
npm run sync -- "변경 내용 요약"
```

이 명령은 다음 순서로 동작합니다.

1. TypeScript 및 Vite 프로덕션 빌드 검사
2. 최신 `origin/main` 확인 및 필요한 경우 rebase
3. 변경 파일 커밋
4. `origin/main` 푸시

`main`에 푸시되면 GitHub Actions가 자동으로 빌드하고 `gh-pages` 브랜치에 배포합니다. 자세한 현장 사용법과 구현 상태는 [사용 가이드](docs/USER_GUIDE.md)와 [프로젝트 상태](docs/PROJECT_STATUS.md)를 참고하세요.
