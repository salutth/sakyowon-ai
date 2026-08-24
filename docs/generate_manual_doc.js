const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, PageNumber } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function cell(text, opts = {}) {
  const { bold, width, shading, align } = opts;
  return new TableCell({
    borders,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: 'center',
    children: [new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: [new TextRun({ text: text || '', bold: !!bold, font: 'Arial', size: 20 })]
    })]
  });
}

function headerCell(text, width) {
  return cell(text, { bold: true, width, shading: 'D5E8F0' });
}

function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function heading3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function para(text, opts = {}) {
  return new Paragraph({ spacing: { after: opts.after || 120 }, alignment: opts.align,
    children: [new TextRun({ text, size: opts.size || 21, bold: opts.bold, color: opts.color, font: 'Arial' })] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 21, font: 'Arial' })] });
}

const rivers = [
  ['도림천', '관악구·동작구·영등포구', '약 14.2km'],
  ['양재천', '서초구·강남구', '약 15.6km'],
  ['중랑천', '도봉구·노원구·성북구·동대문구·성동구', '약 20.2km'],
  ['안양천', '금천구·구로구·영등포구', '약 13.4km'],
  ['탄천', '강남구·송파구', '약 14.0km'],
  ['홍제천', '서대문구·마포구', '약 8.6km'],
  ['불광천', '은평구·서대문구', '약 5.8km'],
  ['정릉천', '성북구', '약 5.3km'],
  ['성북천', '성북구·종로구', '약 4.8km'],
  ['우이천', '강북구·도봉구', '약 6.5km'],
  ['묵동천', '중랑구·노원구', '약 4.2km'],
  ['방학천', '도봉구', '약 5.1km'],
  ['전농천', '동대문구', '약 3.2km'],
  ['면목천', '중랑구', '약 3.8km'],
  ['월계천', '노원구·성북구', '약 4.5km'],
  ['여의천', '영등포구', '약 2.8km'],
  ['반포천', '서초구', '약 3.5km'],
  ['사당천', '동작구', '약 3.1km'],
  ['대림천', '영등포구·구로구', '약 2.5km'],
  ['신월천', '양천구', '약 4.0km'],
  ['구로천', '구로구', '약 3.0km'],
  ['한강(서울)', '강동구 암사동~강서구 개화동', '약 41km (서울 구간)'],
];

const pages = [
  { name: '대시보드 (dashboard.html)', icon: '📊',
    desc: '실시간 수위 현황, AI 홍수 예측, K-SAFE 위험지수를 한눈에 확인할 수 있습니다.',
    features: ['22개 하천 실시간 수위 모니터링', '24시간 수위 추이 차트', 'AI 기반 홍수 예측 (24시간)', 'K-SAFE 위험지수 색상 표시', '최근 홍수 경보 알림'] },
  { name: '오늘의 하천 (today.html)', icon: '📅',
    desc: '오늘 수집된 모든 데이터를 하천별로 요약하여 보여줍니다.',
    features: ['당일 수위/수질 데이터 요약', '하천 선택 필터', '수집 시간대별 정리'] },
  { name: '하천ON 지도 (river-on.html)', icon: '🗺️',
    desc: '문화재, 시민 측정 지점, 관측소를 지도 위에 표시합니다.',
    features: ['Leaflet 기반 인터랙티브 지도', '문화재 마커 표시', '시민 수질 측정 지점', '수위 관측소 위치'] },
  { name: '자동 보고서 (analysis.html)', icon: '📋',
    desc: '수질, 생태건강성(EHI), 트렌드를 자동 분석합니다.',
    features: ['하천별 생태건강성지수(EHI) 등급 표시', '수질 트렌드 분석', '교란종 현황', '전체 하천 비교'] },
  { name: '시민 제보 (report.html)', icon: '🔬',
    desc: '생물 관찰 내용을 사진과 함께 제보합니다.',
    features: ['생물 이름·하천 선택', 'GPS 자동 수집 (위치 권한 필요)', '사진 첨부', '미션 페이지 연동 (체크리스트 자동 갱신)'] },
  { name: '하천 탐험 (play.html)', icon: '🎮',
    desc: '시민 참여형 생물 관찰과 퀴즈를 제공합니다.',
    features: ['하천별 생물 퀴즈', '관찰 미션', '학습 콘텐츠'] },
  { name: '미션 (mission.html)', icon: '🎯',
    desc: '시민과학 미션과 생물 찾기 체크리스트를 제공합니다.',
    features: ['22개 하천별 미션 종 목록', '체크리스트 (localStorage 저장)', 'iNaturalist 종 사진 연동', '📷제보 바로가기 (report.html 연결)'] },
  { name: '거시기 (geosigi)', icon: '🌏',
    desc: '외국인 유학생 정착과 시민과학 참여를 지원합니다.',
    features: ['다국어 지원 (한국어/영어)', '정착 정보', '시민과학 참여 안내'] },
  { name: '플랫폼 현황 (platform.html)', icon: '🏗️',
    desc: '시스템 상태와 수집기 헬스체크 정보를 표시합니다.',
    features: ['데이터 수집기 상태 확인', '마지막 수집 시간', '테이블별 레코드 수'] },
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1a1a2e" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "555555" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1200, bottom: 1440, left: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "RiverWatch v2.3 사용자설명서", font: "Arial", size: 16, color: "999999" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 16, color: "999999" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "999999" })
          ]
        })]
      })
    },
    children: [
      // === Title Page ===
      new Paragraph({ spacing: { before: 2400 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "RiverWatch v2.3", font: "Arial", size: 56, bold: true, color: "7873f5" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
        children: [new TextRun({ text: "사용자 설명서", font: "Arial", size: 40, color: "333333" })] }),
      para("서울하천 AI 모니터링 플랫폼", { align: AlignmentType.CENTER, size: 22, color: "888888" }),
      para("https://dorimchun-ai.pages.dev", { align: AlignmentType.CENTER, size: 22, color: "7873f5" }),
      para("작성일: 2026-07-06", { align: AlignmentType.CENTER, size: 22, color: "888888" }),
      para("도깨비3.0 · AI활동가 1기 · 건강한도림천을만드는주민모임", { align: AlignmentType.CENTER, size: 22, color: "888888" }),

      new Paragraph({ children: [new PageBreak()] }),

      // === 1. 소개 ===
      heading1("1. RiverWatch 소개"),
      para("RiverWatch는 서울특별시 22개 하천의 수위, 수질, 생태건강성을 AI 기반으로 모니터링하는 시민과학 플랫폼입니다."),
      para("주요 목적:", { bold: true }),
      bullet("서울 도시하천의 실시간 수위 모니터링 및 홍수 예측"),
      bullet("시민 참여를 통한 생물 관찰 및 수질 측정 데이터 수집"),
      bullet("생태건강성지수(EHI) 기반 하천 건강 상태 평가"),
      bullet("문화재, 관측소, 생태 정보의 통합 시각화"),

      new Paragraph({ spacing: { before: 200 }, children: [] }),
      heading2("1.1 접속 방법"),
      para("웹 브라우저에서 다음 주소로 접속합니다:"),
      para("https://dorimchun-ai.pages.dev", { bold: true, color: "7873f5" }),
      bullet("PC: Chrome, Firefox, Edge, Safari 등 최신 브라우저"),
      bullet("모바일: iOS Safari, Android Chrome (반응형 지원)"),
      bullet("PWA: 홈 화면에 추가하면 앱처럼 사용 가능"),

      heading2("1.2 Beta 접근"),
      para("현재 Beta 버전으로 비밀번호 입력 후 사용 가능합니다. 비밀번호는 운영팀에 문의해 주세요."),

      new Paragraph({ children: [new PageBreak()] }),

      // === 2. 모니터링 대상 하천 ===
      heading1("2. 모니터링 대상 하천 (22개)"),
      para("서울특별시 관할 22개 하천을 모니터링합니다. 한강은 서울 구간(강동구 암사동~강서구 개화동, 약 41km)으로 한정합니다."),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [600, 2200, 4226, 2000],
        rows: [
          new TableRow({ children: [
            headerCell("No", 600), headerCell("하천명", 2200), headerCell("경유 자치구", 4226), headerCell("연장", 2000)
          ]}),
          ...rivers.map((r, i) => new TableRow({ children: [
            cell(String(i + 1), { width: 600, align: AlignmentType.CENTER }),
            cell(r[0], { width: 2200, bold: true }),
            cell(r[1], { width: 4226 }),
            cell(r[2], { width: 2000 }),
          ]}))
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // === 3. 페이지별 사용법 ===
      heading1("3. 페이지별 사용법"),

      ...pages.flatMap(p => [
        heading2(`${p.icon} ${p.name}`),
        para(p.desc),
        para("주요 기능:", { bold: true }),
        ...p.features.map(f => bullet(f)),
        new Paragraph({ spacing: { after: 160 }, children: [] }),
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // === 4. 시민 제보 가이드 ===
      heading1("4. 시민 제보 가이드"),

      heading2("4.1 생물 관찰 제보"),
      para("미션 페이지에서 제보하기:"),
      bullet("1. 미션(mission.html) 페이지에서 하천과 종류(어류/조류/기타)를 선택합니다"),
      bullet("2. 찾고자 하는 생물 옆의 '📷제보' 링크를 클릭합니다"),
      bullet("3. 제보 페이지에서 생물 이름과 하천이 자동 입력됩니다"),
      bullet("4. 관찰자 이름(선택), 메모, 사진을 추가합니다"),
      bullet("5. GPS 위치가 자동 수집됩니다 (위치 권한 허용 필요)"),
      bullet("6. '제보하기' 버튼을 눌러 제출합니다"),
      bullet("7. 제보 성공 시 미션 체크리스트에 자동으로 체크됩니다"),

      heading2("4.2 시민 수질 측정"),
      para("대시보드의 시민 수질 측정 탭에서:"),
      bullet("1. 하천을 선택합니다 (22개 하천)"),
      bullet("2. 측정 항목을 입력합니다: pH, 용존산소(DO), 수온"),
      bullet("3. 관찰자 이름(선택)을 입력합니다"),
      bullet("4. '기록하기' 버튼으로 제출합니다"),

      heading2("4.3 제보 시 주의사항"),
      bullet("GPS 위치 권한을 허용해야 정확한 관찰 위치가 기록됩니다"),
      bullet("생물 이름은 정확한 국명 또는 학명을 사용해 주세요"),
      bullet("사진은 선명하게, 생물의 특징이 잘 보이도록 촬영해 주세요"),
      bullet("허위 제보는 데이터 품질을 저해합니다"),

      new Paragraph({ children: [new PageBreak()] }),

      // === 5. 생태건강성지수(EHI) ===
      heading1("5. 생태건강성지수(EHI) 이해하기"),
      para("EHI(Ecosystem Health Index)는 하천의 생태 건강 상태를 0~100점으로 평가하는 종합 지수입니다."),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1400, 6226],
        rows: [
          new TableRow({ children: [headerCell("등급", 1400), headerCell("점수", 1400), headerCell("의미", 6226)] }),
          new TableRow({ children: [cell("A", { width: 1400, align: AlignmentType.CENTER, shading: "C8E6C9" }), cell("80~100", { width: 1400 }), cell("매우 양호: 생태계가 건강하고 다양한 종이 서식", { width: 6226 })] }),
          new TableRow({ children: [cell("B", { width: 1400, align: AlignmentType.CENTER, shading: "DCEDC8" }), cell("60~79", { width: 1400 }), cell("양호: 전반적으로 건강하나 일부 스트레스 요인 존재", { width: 6226 })] }),
          new TableRow({ children: [cell("C", { width: 1400, align: AlignmentType.CENTER, shading: "FFF9C4" }), cell("40~59", { width: 1400 }), cell("보통: 생태계 스트레스 징후, 관리 필요", { width: 6226 })] }),
          new TableRow({ children: [cell("D", { width: 1400, align: AlignmentType.CENTER, shading: "FFE0B2" }), cell("20~39", { width: 1400 }), cell("나쁨: 생태계 기능 저하, 적극적 복원 필요", { width: 6226 })] }),
          new TableRow({ children: [cell("F", { width: 1400, align: AlignmentType.CENTER, shading: "FFCDD2" }), cell("0~19", { width: 1400 }), cell("매우 나쁨: 생태계 심각한 훼손", { width: 6226 })] }),
        ]
      }),

      para("EHI 산출 항목:", { bold: true, after: 80 }),
      bullet("생물다양성 점수 (biodiversity_score): 관찰 종 수 기반"),
      bullet("수위안정성 점수 (water_stability_score): 수위 변동폭"),
      bullet("비교란종 점수 (non_invasive_score): 외래종 비율 역산"),
      bullet("관찰빈도 점수 (observation_freq_score): 시민 관찰 활동성"),

      new Paragraph({ children: [new PageBreak()] }),

      // === 6. 데이터 수집 체계 ===
      heading1("6. 데이터 수집 체계"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 2200, 2200, 2426],
        rows: [
          new TableRow({ children: [headerCell("수집기", 2200), headerCell("데이터", 2200), headerCell("주기", 2200), headerCell("소스", 2426)] }),
          new TableRow({ children: [cell("species.py", { width: 2200 }), cell("종 관찰 데이터", { width: 2200 }), cell("매일 1회", { width: 2200 }), cell("iNaturalist API", { width: 2426 })] }),
          new TableRow({ children: [cell("weather.py", { width: 2200 }), cell("기상·수위 데이터", { width: 2200 }), cell("매시간", { width: 2200 }), cell("서울시 Open API", { width: 2426 })] }),
          new TableRow({ children: [cell("ehi.py", { width: 2200 }), cell("생태건강성지수", { width: 2200 }), cell("매일 1회", { width: 2200 }), cell("자체 산출", { width: 2426 })] }),
          new TableRow({ children: [cell("cultural_assets.py", { width: 2200 }), cell("문화재 정보", { width: 2200 }), cell("주 1회", { width: 2200 }), cell("문화재청 API", { width: 2426 })] }),
        ]
      }),

      para("모든 수집기는 GitHub Actions를 통해 자동 실행되며, Supabase에 저장됩니다.", { after: 200 }),

      // === 7. 보안 관련 안내 ===
      heading1("7. 보안 관련 안내"),

      heading2("7.1 개인정보 처리"),
      bullet("제보 시 수집 정보: 닉네임(선택), GPS 좌표, 종 이름, 하천, 메모"),
      bullet("여권번호, 비자번호, 외국인등록번호 등 민감한 개인정보는 수집하지 않습니다"),
      bullet("GPS 위치는 생물 관찰 지점 확인 목적으로만 사용됩니다"),

      heading2("7.2 데이터 보안"),
      bullet("모든 데이터는 Supabase (PostgreSQL) 클라우드에 안전하게 저장됩니다"),
      bullet("RLS(Row Level Security) 정책으로 데이터 접근이 제어됩니다"),
      bullet("HTTPS 암호화 통신만 사용합니다"),
      bullet("Cloudflare CDN을 통한 DDoS 방어가 적용되어 있습니다"),

      heading2("7.3 사용자 유의사항"),
      bullet("Beta 비밀번호를 타인에게 공유하지 마세요"),
      bullet("제보 데이터는 공개됩니다 (민감한 개인정보를 메모에 입력하지 마세요)"),
      bullet("허위 데이터 제보는 시민과학의 신뢰성을 훼손합니다"),

      new Paragraph({ children: [new PageBreak()] }),

      // === 8. FAQ ===
      heading1("8. 자주 묻는 질문 (FAQ)"),

      heading3("Q. 한강도 모니터링하나요?"),
      para("네. 한강(서울) 구간, 즉 서울특별시 내 한강 본류(강동구 암사동 ~ 강서구 개화동, 약 41km)를 모니터링합니다."),

      heading3("Q. 데이터는 얼마나 자주 업데이트되나요?"),
      para("수위/기상 데이터는 매시간, 종 관찰 데이터는 매일, 생태건강성지수(EHI)는 매일 갱신됩니다."),

      heading3("Q. 제보한 생물이 미션 체크리스트에 반영되지 않아요."),
      para("제보 성공 시 자동으로 체크됩니다. 반영되지 않으면 브라우저의 localStorage가 가득 찼거나 시크릿 모드일 수 있습니다."),

      heading3("Q. 거시기(Geosigi)는 무엇인가요?"),
      para("외국인 유학생의 한국 정착과 시민과학 참여를 지원하는 다국어 앱입니다. 현재 개발 중입니다."),

      heading3("Q. 오프라인에서도 사용할 수 있나요?"),
      para("서비스 워커(PWA)가 적용되어 있어, 이전에 방문한 페이지는 오프라인에서도 볼 수 있습니다. 단, 실시간 데이터는 인터넷 연결이 필요합니다."),

      heading3("Q. 버그를 발견하면 어떻게 신고하나요?"),
      para("GitHub Issues에 보고해 주세요: https://github.com/salutth/dorimchun-ai/issues"),

      new Paragraph({ spacing: { before: 600 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "--- 문서 끝 ---", color: "999999", size: 20 })
      ] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2] || 'RiverWatch_User_Manual.docx', buf);
  console.log('User manual generated');
});
