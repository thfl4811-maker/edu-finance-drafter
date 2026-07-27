# 품의서 작성기 AI 전면교체본

## 핵심 기능
- 엑셀·PDF·이미지 견적서를 AI가 문맥으로 분석
- 합계·총계·소계·부가세 행 제외
- 부가세 포함 예상단가 계산
- 사용자가 품목을 확인·수정
- 사용자가 제공한 `품목내역(통합) (2).xls` 원본의 4개 열에만 입력
  - 내용
  - 규격
  - 수량
  - 예상단가
- 물품 구입 품의서와 업무추진비·협의회 품의서 작성

## Vercel 환경변수
반드시 다음 환경변수를 등록합니다.

- `OPENAI_API_KEY`: OpenAI API 키
- `OPENAI_MODEL`: 선택 사항. 기본값 `gpt-4.1-mini`

## 배포
1. 압축을 풀고 파일 전체를 GitHub 저장소 루트에 업로드합니다.
2. 기존 저장소 파일은 이 배포본으로 교체합니다.
3. Vercel Framework Preset은 `Vite`, Root Directory는 `./`로 둡니다.
4. 환경변수를 등록한 후 Deploy 합니다.


## 압축파일 검증

이 배포본에는 0바이트 일반 파일이 없습니다.

- `api/analyze.js`: AI 견적서 분석 API
- `src/main.js`: 화면, 파일 분석, 에듀파인 엑셀 생성
- `src/style.css`: 화면 디자인
- `public/edufine-template.xls`: 에듀파인 원본 서식
- `index.html`, `package.json`, `vercel.json`: 배포 설정

압축 목록에서 `api/`, `public/`, `src/`가 0바이트로 보이는 경우는 파일이 아니라 폴더 항목입니다.
이번 검토본은 혼동을 막기 위해 폴더 항목을 별도로 넣지 않고 실제 파일만 압축했습니다.
