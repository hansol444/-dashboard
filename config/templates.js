function today() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return { display: `${y}.${m}.${day}`, yyyymm: `${y}${m}` };
}

function getBudgetClass(total) {
  if (total <= 1000000)  return '┗━━5)예산 배정(신규, 전용, 초과) - 100만원 이하';
  if (total <= 10000000) return '┗━━4)예산 배정(신규, 전용, 초과) - 100~1,000만원 이하';
  return                        '┗━━3)예산 배정(신규, 전용, 초과) - 1,000만원 초과';
}

function getPurchaseClass(total) {
  if (total < 1000000)   return '┗━ 1) 100만원 미만';
  if (total < 10000000)  return '┗━ 2) 100만원 이상 ~ 1,000만원 미만';
  return                        '┗━ 3) 1,000만원 이상';
}

module.exports = {

  'ats-b': (args) => {
    const t = today();
    const nq      = parseInt(args.nq      ?? 12);     // 네이버페이 수량
    const sq      = parseInt(args.sq      ?? 10);     // 스타벅스 수량
    const nPrice  = parseInt(args.nprice  ?? 30000);  // 네이버페이 단가 (기본 3만원)
    const sPrice  = parseInt(args.sprice  ?? 10000);  // 스타벅스 단가 (기본 1만원)
    const disc    = parseFloat(args.disc  ?? 0.98);   // 할인율 (기본 2% 할인)
    const mon     = args.mon ?? t.yyyymm;
    const si      = args.si  ?? '';
    const na      = Math.round(nPrice * nq * disc);
    const sa      = Math.round(sPrice * sq * disc);
    const tot     = na + sa;
    const totMan  = Math.round(tot / 10000);
    const nPriceMan = Math.round(nPrice / 10000);
    const sPriceMan = Math.round(sPrice / 10000);
    return {
      기본: {
        문건제목: '[예산품의] ATS 인터뷰 진행에 따른 기프티콘 구매를 위한 예산품의',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getBudgetClass(tot),
        기안부서: 'CEO(전략추진실)',
        기안자: '이창준',
        시행일: si,
      },
      예산연동: [
        { 구분: 'LE확인', from: { 신청월: mon, 코스트센터: 'D2034', GL계정: '343006', 금액: tot }, to: { 신청월: mon, 코스트센터: 'D2034', GL계정: '343006', 금액: tot } }
      ],
      본문: `1. 목적

ATS 사업 관련 전략 수립을 위한 인사이트 확보를 목적으로, 현직 인사담당자를 대상으로 인터뷰를 진행하였음.
인터뷰에 참여해주신 분들께 감사의 뜻을 전하고, 향후 재협업 가능성을 높이기 위한 관계 유지 차원에서 소정의 기프티콘(네이버페이)을 제공하고자 함.
해당 인터뷰는 잡코리아 및 나인하이어 관련 사업 전략 수립에 직접적으로 활용될 예정임.

2. 세부 추진 계획

인터뷰 대상자 수: 총 ${nq + sq}명
기프티콘 종류: 3만원 상당 네이버페이 상품권, 1만원 상당 스타벅스 상품권
지급 시점: 인터뷰 완료 후 개별 문자 발송 (발송 기록 확보 예정)
기프티콘 구매처: 공식 기프티콘 대량 발송 플랫폼 활용 (KT 알파 2% 할인 적용)

3. 예산 상세 내역

- 항목: 네이버페이 상품권 (3만원권)
- 단가: ${nPrice.toLocaleString()}원 / 수량: ${nq}건 / 합계: ${(nPrice*nq).toLocaleString()}원
- 총예산: ${na.toLocaleString()}원 (2% 할인 적용)
- 계정과목: 광고선전비(컨텐츠)

- 항목: 스타벅스 상품권 (1만원권)
- 단가: ${sPrice.toLocaleString()}원 / 수량: ${sq}건 / 합계: ${(sPrice*sq).toLocaleString()}원
- 총예산: ${sa.toLocaleString()}원 (2% 할인 적용)
- 계정과목: 광고선전비(컨텐츠)

4. 산출 근거

인터뷰 대상자 1인당 3만원, 1만원 기준
실제 인터뷰 완료 건 기준으로만 발송하며, 기록 보관을 통해 집행 증빙 확보

5. 기대효과 및 목표

ATS 사업 관련 니즈 및 현업 활용도에 대한 생생한 인사이트 확보
자사 서비스(잡코리아/나인하이어)에 대한 고객 관점 피드백 수집
인터뷰 대상자와의 긍정적 관계 형성을 통한 향후 재접점 기반 마련`,
    };
  },

  'ats-p': (args) => {
    const t = today();
    const nq = parseInt(args.nq ?? 12);
    const sq = parseInt(args.sq ?? 10);
    const na  = Math.round(30000 * nq * 0.98);
    const sa  = Math.round(10000 * sq * 0.98);
    const tot = na + sa;
    return {
      기본: {
        문건제목: '[구매품의] ATS 인터뷰 진행에 따른 기프티콘 구매',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getPurchaseClass(tot),
        기안부서: 'CEO(전략추진실)',
        기안자: '이창준',
        시행일: t.display,
      },
      구매품의: {
        요청유형: '비용',
        계약유무: '아니오',
        공급사: 'KT 알파 (기프티콘 대량발송 플랫폼)',
        계약기간: '',
        품목: [
          { 품목명: `네이버페이 상품권 3만원권 (KT알파 2% 할인)`, 수량: nq, 단위: '건', 단가: 30000, 비용: na },
          { 품목명: `스타벅스 상품권 1만원권 (KT알파 2% 할인)`,  수량: sq, 단위: '건', 단가: 10000, 비용: sa },
        ],
        총비용: tot,
      },
    };
  },

  'survey-b': (args) => {
    const t  = today();
    const yr = args.yr ?? new Date().getFullYear().toString();
    const amt = parseInt(args.amt ?? 15000000);

    let months;
    if (args.mon) {
      months = [String(args.mon)];
    } else if (args.months) {
      months = String(args.months).split(',').map(m => m.trim());
    } else {
      months = [`${yr}03`, `${yr}06`, `${yr}09`, `${yr}12`];
    }

    const total = amt * months.length;
    const totalMan = Math.round(total / 10000);
    const amtMan = Math.round(amt / 10000);

    return {
      기본: {
        문건제목: '[예산품의] Placement Survey 예산 집행',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getBudgetClass(total),
        기안부서: 'CEO(전략추진실)',
        기안자: '주호연',
        시행일: '',
      },
      예산연동: months.map(mon => ({
        구분: 'LE확인',
        from: { 신청월: mon, 코스트센터: 'M2010', GL계정: '343002', 금액: amt },
        to:   { 신청월: mon, 코스트센터: 'M2010', GL계정: '343002', 금액: amt },
      })),
      본문: `1. 배경 및 목적

채용 시장은 그 특성 상 시장 규모가 거시 경제(경제 성장률, 실업률 등)에 따라 쉽게 변동되는 특성을 지니고 있습니다.
따라서 단순히 '매출'의 증가를 회사가 '더 건강하고 경쟁력 있는 좋은 회사'가 된 증거로 보기 어렵습니다.
실제로 2022년 COVID 때 국내 채용업계는 굉장히 큰 재무적 성과를 얻었지만, 그 당시 대비 시장 규모는 20~30% 줄어들었습니다.
이에 따라 회사가 정말 더 많은 사람들을 채용하는데 도움을 주고 있는지를 알아내기 위해 Survey를 진행해왔고, 그 Survey가 'Job Placement Survey'입니다.
기존 방식의 경우 표본에 있어서 제대로 된 통제가 없다보니 국내 현실을 제대로 대변하고 있는지 의구심이 있었습니다.
회사 차원에서도 North Star 관점에서 Placement Share%를 설정한 만큼, Survey의 품질을 올리기 위해 엠브레인 측과 논의하여 표본 조건을 강화하고 질문지를 수정하였으며, 조건이 타이트해진 Survey를 올해도 지속적으로 유지하고 있습니다.

2. 필요 예산 및 방식

총액: ${totalMan}만원
${months.map(m => `- ${m}: ${amtMan}만원`).join('\n')}
- 코스트센터: M2010 잡코리아마케팅팀
- G/L계정: 343002 광고선전비(조사연구)
결제 방식: 법인카드 결제
상세 내용 견적서 첨부파일 올렸습니다.

3. 예산의 기대 효과

저희 주주사인 SEEK에서 하는 방식 대비 아직 고도화되진 않았지만, 기존 방식 대비 높은 정확도의 Placement Share% 측정을 통해 우리가 진짜 성장하고 있는지 제대로 측정할 수 있습니다.
기존에 하기 어려웠던 '세그먼트별 Share', 즉 지역, 연령, 연봉, 직무, 회사 특성별 경쟁사 대비 어떠한 위치에 있는지 알아냄으로써 전사 경영 의사결정에 기여할 것으로 기대합니다.`,
    };
  },

  'survey-p': (args) => {
    const t  = today();
    const yr = args.yr ?? new Date().getFullYear().toString();
    const unitAmt = parseInt(args.amt ?? 14000000);
    const total = unitAmt * 4;
    return {
      기본: {
        문건제목: '[구매품의] Placement Survey',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getPurchaseClass(total),
        기안부서: 'CEO(전략추진실)',
        기안자: '주호연',
        시행일: t.display,
      },
      구매품의: {
        요청유형: '비용',
        계약유무: '예',
        공급사: '(주) 마크로밀 엠브레인',
        계약기간: `${yr}년 1월 1일 ~ ${yr}년 12월 31일`,
        품목: [
          { 품목명: '1~3월 (1분기) 조사용역비',   수량: 1, 단위: 'ea', 단가: unitAmt, 비용: unitAmt },
          { 품목명: '4~6월 (2분기) 조사용역비',   수량: 1, 단위: 'ea', 단가: unitAmt, 비용: unitAmt },
          { 품목명: '7~9월 (3분기) 조사용역비',   수량: 1, 단위: 'ea', 단가: unitAmt, 비용: unitAmt },
          { 품목명: '10~12월 (4분기) 조사용역비', 수량: 1, 단위: 'ea', 단가: unitAmt, 비용: unitAmt },
        ],
        총비용: total,
      },
    };
  },

  'free-b': (args) => {
    const t     = today();
    const mon   = args.mon   ?? t.yyyymm;
    const n     = parseInt(args.n     ?? 9);      // 인원수 (기본 9명)
    const weeks = parseInt(args.weeks ?? 6);      // 진행 주차 (기본 6주)
    const hpw   = parseInt(args.hpw   ?? 40);     // 1주당 시간 (기본 40시간)
    const totalHours = n * weeks * hpw;           // 총 시간
    const hourlyWage = Math.round(10000);         // 시급 약 1만원
    const amt   = parseInt(args.amt ?? (totalHours * hourlyWage)); // 총액 자동계산
    const amtMan = Math.round(amt / 10000);
    const hourly = Math.round(amt / totalHours);
    return {
      기본: {
        문건제목: '[예산품의] 전략추진실 - 산학협력 프리랜서 ${weeks}주 고용',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getBudgetClass(amt),
        기안부서: 'CEO(전략추진실)',
        기안자: '임성욱',
        시행일: t.display,
      },
      본문: `1. 배경 및 목적

전략추진실의 정규직 인력 TO가 5명이나 현재 0명 상태로 운영되고 있고, 인턴분들을 통해 일 진행하고 있지만 현안 과제가 많아 업무 부하가 높은 상황입니다.
이를 산학협력 형태의 경영전략학회 대학생 프리랜서 고용을 통해 일부 해결하고자 합니다.
특히, 산학협력 형태를 통할 경우 최저임금보다 낮은 비용으로 일을 같이 할 수 있어 (= 업무 경험 체험의 의미) 비용 효율적일 것으로 판단됩니다.
이번 건의 경우 인수하게 된 잡플래닛의 향후 전략에 대해서 Worxphere 본사 차원에서 전략 방향 고민 진행 예정입니다 (의사결정/커뮤니티)

2. 필요 예산 및 방식

1회 ${amtMan}만원 (VAT 포함)
- ${n}명 인원, ${weeks}주, 1주 ${hpw}시간 업무
- 총 ${totalHours}시간 업무로 시급 환산 시 약 ${hourly.toLocaleString()}원 수준
- 프리랜서 고용하는 계약서 통해서 대표자 1인을 프리랜서 고용하는 것으로 하여 용역 대금 지급하는 구조
2026년 사업계획 상에 1,800만원 예산이 배정되었으며, 해당 예산을 사용하겠습니다.

3. 예산의 기대 효과

회사의 여러 전략 아젠다에 대해서 최대한 퀄리티를 낮추지 않고 많은 영역을 커버하여 회사의 성과를 높이는 데 기여할 것으로 기대합니다.`,
    };
  },

  'free-p': (args) => {
    const t   = today();
    const sup = args.sup ?? 'EGI';
    const st  = args.st  ?? '';
    const en  = args.en  ?? '';
    const amt = parseInt(args.amt ?? 5000000);
    return {
      기본: {
        문건제목: '[구매품의] 산학협력 프리랜서 6주 고용',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getPurchaseClass(amt),
        기안부서: 'CEO(전략추진실)',
        기안자: '임성욱',
        시행일: '',
      },
      구매품의: {
        요청유형: '비용',
        계약유무: '예',
        공급사: sup,
        계약기간: `${st} ~ ${en}`,
        품목: [
          { 품목명: '산학협력 프리랜서 6주 고용', 수량: null, 단위: null, 단가: null, 비용: amt },
        ],
        총비용: amt,
      },
    };
  },


  // ─────────────────────────────────────────────
  // 산학협력 인터뷰 비용 지급 — 예산품의
  // 사용법: node fill.js interview-b --mon 202603 --from_mon 202612 --n30 40 --n60 20
  // ─────────────────────────────────────────────
  'interview-b': (args) => {
    const t   = today();
    const toMon   = args.mon      ?? t.yyyymm;        // TO 신청월 (집행월)
    const fromMon = args.from_mon ?? toMon;            // FROM 신청월 (차감월)
    const n30    = parseInt(args.n30    ?? 40);      // 30분 인터뷰 인원
    const n60    = parseInt(args.n60    ?? 20);      // 1시간 인터뷰 인원
    const rate30 = parseInt(args.rate30 ?? 30000);   // 30분 단가 (기본 3만원)
    const rate60 = parseInt(args.rate60 ?? 50000);   // 1시간 단가 (기본 5만원)
    const amt    = parseInt(args.amt ?? (n30 * rate30 + n60 * rate60)); // 총액 자동계산
    const amtMan = Math.round(amt / 10000);
    const rate30Man = Math.round(rate30 / 10000);
    const rate60Man = Math.round(rate60 / 10000);

    return {
      기본: {
        문건제목: '[예산품의] 전략추진실 산학협력 인터뷰 비용 지급을 위한 예산품의',
        읽기권한: '결재권자 + 동일팀',
        문건분류선택: getBudgetClass(amt),
        기안부서: 'CEO(전략추진실)',
        기안자: '주호연',
        시행일: t.display,
      },
      예산연동: [
        {
          구분: args.type_sap ?? '전용',
          from: { 신청월: fromMon, 코스트센터: args.from_cc ?? 'M2010', GL계정: '343002', 금액: amt },
          to:   { 신청월: toMon,   코스트센터: args.to_cc   ?? 'H2200', GL계정: '343002', 금액: amt },
        }
      ],
      본문: `1. 배경 및 목적

전략추진실의 정규직 인력 TO가 5명이나 현재 0명 상태로 운영되고 있고, 인턴분들을 통해 일 진행하고 있지만 현안 과제가 많아 업무 부하가 높은 상황에 진행하고 있던 산학협력에서 구직자 인터뷰가 필요합니다.
이번 건의 경우 인수하게 된 잡플래닛의 향후 전략에 대해서 Worxphere 본사 차원에서 전략 방향 고민 진행 중이며, 관련하여 응해준 구직자 대상으로 인터뷰 보상을 지급하고자 합니다.

2. 필요 예산 및 방식

30분 기준 ${rate30Man}만원, 1시간 기준 ${rate60Man}만원.
총 ${n30 + n60}명 진행
30분 ${n30}명, 1시간 ${n60}명 진행 시 ${amtMan}만원 소요
잡플레이스먼트 계약 당시 배정된 금액에 비해 남은 금액을 활용하겠습니다.

3. 예산의 기대 효과

회사의 여러 전략 아젠다에 대해서 최대한 퀄리티를 낮추지 않고 많은 영역을 커버하여 회사의 성과를 높이는 데 기여할 것으로 기대합니다.`,
    };
  },

};
