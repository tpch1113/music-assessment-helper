# 음악 수행평가 채점 도우미

React + Vite 기반의 중학교 음악 수행평가 채점 웹앱입니다.

## 실행 방법

```bash
npm install
npm run dev
```

브라우저에서 안내되는 로컬 주소로 접속하면 사용할 수 있습니다.

## 주요 기능

- 수행평가명 입력
- 평가영역 추가, 수정, 삭제
- 평가영역별 배점 입력
- 세부 채점 기준 및 점수 구간 관리
- 채점 기준표 저장 및 불러오기
- 학생 목록 일괄 입력
- 학생 리스트에서 학생 선택 후 바로 채점
- 완료/미채점 상태 표시
- 저장 후 자동으로 다음 학생으로 이동
- 이전 학생 / 다음 학생 이동
- 미채점 학생만 보기 및 완료 학생 숨기기
- 숫자 키 1~5로 빠른 점수 입력
- 학생 작품 텍스트와 평가기준을 비교한 AI 점수 추천
- 학생 작품 PDF 업로드 및 텍스트 추출
- 학생 작품 사진 여러 장 업로드 및 OpenAI Vision 분석
- 여러 학생 작품 사진 일괄 업로드 및 학생 자동 매칭
- 매칭된 학생의 작품 사진 있음 표시
- OpenAI API Key localStorage 저장
- 학생별 점수 선택 및 영역 배점 기준 총점 자동 계산
- 선택 점수와 교사 메모 기반의 `~함.` 문체 피드백 자동 생성
- 채점 결과 저장, 수정, 삭제
- 전체 학생 결과 CSV 다운로드
- localStorage 자동 저장
- 아이패드와 노트북에 맞춘 반응형 화면

## AI 채점 보조 사용 방법

1. `설정` 탭에서 OpenAI API Key를 입력합니다.
2. `학생 목록` 또는 `채점` 탭에서 채점할 학생을 선택합니다.
3. `AI 보조` 탭의 `여러 학생 사진 일괄 업로드`에서 사진을 한 번에 선택합니다.
4. 파일명은 `1-01-김민서.jpg`, `1_02_박지훈.png`, `3301_이서연.jpg`처럼 반, 번호, 이름이 보이게 준비합니다.
5. 학생 목록과 매칭된 학생은 `작품 사진 있음` 표시가 나오며, 학생을 클릭하면 해당 사진들이 AI 보조 탭에 자동으로 표시됩니다.
6. 매칭되지 않은 파일은 `매칭 실패 파일` 목록에서 확인합니다.
7. 필요하면 학생 작품 사진을 직접 추가하거나, PDF 파일을 업로드하거나, 학생 작품 텍스트를 직접 붙여넣습니다.
8. 사진을 업로드하면 AI가 사진 속 글과 작품 내용을 읽고 현재 평가기준과 비교합니다.
9. PDF 텍스트 추출에 실패하면 안내 문구에 따라 내용을 직접 복사해 넣습니다.
10. `AI 점수 추천`을 누르면 평가기준별 추천 점수와 이유가 표시됩니다.
11. `점수 적용`과 `피드백 적용`을 누른 뒤, `채점` 탭에서 교사가 최종 점수를 확인하고 수정합니다.
12. `저장 후 다음`을 눌러 결과를 저장합니다.

## Vercel 배포 방법

### 방법 1. GitHub 없이 Vercel CLI로 바로 배포

1. Node.js가 설치되어 있는지 확인합니다.

```bash
node -v
npm -v
```

2. 프로젝트 폴더에서 의존성을 설치합니다.

```bash
npm install
```

3. 배포 전 빌드가 되는지 확인합니다.

```bash
npm run build
```

4. Vercel CLI를 실행합니다.

```bash
npx vercel
```

5. 처음 실행하면 Vercel 로그인을 진행합니다.
6. 물어보는 항목은 보통 기본값으로 진행하면 됩니다.
   - Framework Preset: `Vite`
   - Build Command: `npm run build`
   - Output Directory: `dist`
7. 실제 공개 배포는 다음 명령으로 실행합니다.

```bash
npx vercel --prod
```

### 방법 2. GitHub에 올린 뒤 Vercel에 연결

1. GitHub에서 새 저장소를 만듭니다.
2. 프로젝트 폴더에서 Git을 초기화하고 파일을 올립니다.

```bash
git init
git add .
git commit -m "Initial deploy"
git branch -M main
git remote add origin https://github.com/사용자명/저장소명.git
git push -u origin main
```

3. Vercel에 접속해 `Add New Project`를 누릅니다.
4. GitHub 저장소를 선택합니다.
5. 설정은 다음처럼 확인합니다.
   - Framework Preset: `Vite`
   - Install Command: `npm install`
   - Build Command: `npm run build`
   - Output Directory: `dist`
6. `Deploy`를 누르면 배포됩니다.

### 배포 후 주의

- OpenAI API Key는 앱의 `설정` 탭에서 사용자 브라우저 localStorage에 저장됩니다.
- API Key를 코드나 GitHub 저장소에 직접 넣지 마세요.
- 공용 PC에서는 사용 후 브라우저 저장 데이터를 삭제하는 것이 좋습니다.
