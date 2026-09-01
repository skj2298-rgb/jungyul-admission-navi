// data/universities/gachon.js
// 가천대학교 — 대학 구조 DB(UNIVERSITY_DB) 등록 파일.
// 사용자가 직접 제공한 가천대 공식 모집요강·입시결과 원본 파일을 근거로 입력했다(추정·임의값 없음).
// 원본 파일(사용자 제공, 다운로드 폴더):
//   2024가천대-수시모집요강_수정(0905).pdf, 2025학년도_가천대학교_수시모집요강.pdf, 2026가천대학교수시모집요강.pdf
//   2024학년도정시모집요강.hwp(파싱 불가), 2026가천대학교정시모집요강수정(단면)_업로드.pdf
//   2024학년도정시입시결과.xls, 2025학년도정시입시결과.xls(비밀번호 보호 — 열지 않음), 2026정시입시결과.xls
//   가천대학교_2024학년도_수시_입시결과.xlsx, 2025학년도수시입시결과.xlsx, 2026학년도수시입시결과.xlsx
//
// 파일럿 범위: 경영학과, 컴퓨터공학과 2개 모집단위 (2024~2026학년도 정시 일반전형 + 수시 학생부우수자전형=학생부교과).
//
// 중요: 아래 계산규칙(CALCULATION_RULES)은 전부 status:'draft'로만 등록했다(=시뮬레이터에 연결되지 않음, "계산 보류").
// 이유는 department.simulationBinding을 모두 null로 둔 것과 함께 각 규칙 notes에 구체적으로 적어 두었다.
// 대표적으로: (1) 정시는 해마다 "국/수/영(/탐) 중 우수한 영역순" 가중치를 적용하는 순위기반 반영방식이라
// 기존 앱의 고정비율 가중합산 엔진(scoreStudent)과 계산 방식 자체가 달라 그대로 재사용할 수 없고, 영어 등급 환산표도
// 이번 자료에서 확인하지 못했다. (2) 학생부교과(학생부우수자)는 2026학년도만 등급→배점 환산표를 완전히 확인했지만,
// 입시결과 파일에 공개된 "학생부등급(70%/90%)" 컷은 원 석차등급 평균(1~9등급 척도)이고 실제 반영 산식의 환산점수
// (70~100점 척도)와 척도가 달라, 서로 다른 단위를 비교하는 오류를 피하기 위해 연결하지 않았다.
(function(){
  const SRC = {org:'가천대학교 입학처(사용자 제공 자료)', url:null, publishedDate:null, checkedDate:'2026-09-01'};

  // ===== 계산규칙(모두 draft = 시뮬레이터 미연결, 원자료·반영방법 문서화 목적) =====
  CALCULATION_RULES['gachon-regular-general1-2024']={
    id:'gachon-regular-general1-2024',university:'gachon',admissionType:'regular',appliesToYears:[2024],status:'draft',
    notes:['2024학년도 정시모집요강이 .hwp 형식으로만 제공되어 자동 파싱이 불가능했다. 반영비율·영어 환산표를 확인하지 못해 draft로 둔다. 모집인원·경쟁률·70%컷 등 결과 수치만 원자료 그대로 등록했다.'],
    sources:[{...SRC,title:'2024학년도정시모집요강.hwp(파싱 불가)'},{...SRC,title:'2024학년도정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-regular-general1-2026']={
    id:'gachon-regular-general1-2026',university:'gachon',admissionType:'regular',appliesToYears:[2026],status:'draft',
    reflectedSubjects:['국어','수학','영어','탐구'],
    method:'인문·자연계열(의예과·한의예과·약학과·클라우드공학과 제외) 공통: 국어·수학·영어 중 성적이 좋은 순서로 35%·25%·20% 반영 + 탐구(1과목) 20% 반영. 활용지표는 백분위.',
    notes:['"우수한 영역순" 순위 기반 가중치 방식이라 기존 앱의 고정비율 가중합산 엔진(scoreStudent)을 그대로 쓸 수 없다.','영어 등급→점수 환산표와 탐구 1과목 선정 기준(2과목 중 어느 과목인지)은 이번 자료에서 확인하지 못했다.'],
    sources:[{...SRC,title:'2026가천대학교정시모집요강수정(단면)_업로드.pdf',page:31},{...SRC,title:'2026정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-regular-general2-2026']={
    id:'gachon-regular-general2-2026',university:'gachon',admissionType:'regular',appliesToYears:[2026],status:'draft',
    reflectedSubjects:['국어','수학','영어','탐구'],
    method:'2026학년도 신설 전형. 국어·수학·영어·탐구(1과목) 중 성적이 좋은 순서로 50%·30%·20%만 반영(나머지 1개 영역은 미반영). 활용지표는 표준점수(영어는 변환표준점수). 2027학년도부터는 활용지표가 등급으로 변경 예정(모집요강 안내).',
    notes:['"우수 3개 영역만" 반영하는 방식이라 기존 앱의 고정비율 4영역 가중합산 엔진을 그대로 쓸 수 없다.'],
    sources:[{...SRC,title:'2026가천대학교정시모집요강수정(단면)_업로드.pdf',page:31},{...SRC,title:'2026정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2024-humanities']={
    id:'gachon-schoolrecord-2024-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',
    reflectedSubjects:['국어','수학','영어','사회'],
    method:'인문계열: 국어·수학·영어·사회 중 우수한 교과 순으로 40%·30%·20%·10% 반영(진로선택과목 반영 없음). 활용지표는 석차등급.',
    notes:['교과별 순위 가중치 방식이며, 등급을 점수로 환산하는 배점표는 이 자료에서 확인하지 못했다(석차등급을 직접 가중평균하는지, 별도 배점표가 있는지 불확실).'],
    sources:[{...SRC,title:'2024가천대-수시모집요강_수정(0905).pdf',page:84},{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2024-science']={
    id:'gachon-schoolrecord-2024-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',
    reflectedSubjects:['국어','수학','영어','과학'],
    method:'자연계열: 국어·수학·영어·과학 중 우수한 교과 순으로 40%·30%·20%·10% 반영(진로선택과목 반영 없음). 활용지표는 석차등급.',
    notes:['교과별 순위 가중치 방식이며, 등급을 점수로 환산하는 배점표는 이 자료에서 확인하지 못했다.'],
    sources:[{...SRC,title:'2024가천대-수시모집요강_수정(0905).pdf',page:84},{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2025-humanities']={
    id:'gachon-schoolrecord-2025-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',
    reflectedSubjects:['국어','수학','영어','사회'],
    method:'인문계열: 국어·수학·영어·사회, 진로선택과목 반영 없음. 요소별 반영비율은 "교과(일반/공통선택과목) 100%"(비교과 없음). 학년-학기별 가중치 없음(3학년 1학기까지 반영). 활용지표는 석차등급.',
    notes:['교과 요소가 100%라는 것만 확인했고, 과목 간/학년 간 세부 가중 방식(단순 이수단위 평균인지 별도 배점표가 있는지)은 이 자료에서 명확히 확인하지 못했다.'],
    sources:[{...SRC,title:'2025학년도_가천대학교_수시모집요강.pdf',page:84}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2025-science']={
    id:'gachon-schoolrecord-2025-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',
    reflectedSubjects:['국어','수학','영어','과학'],
    method:'자연계열: 국어·수학·영어·과학, 진로선택과목 반영 없음. 요소별 반영비율은 "교과(일반/공통선택과목) 100%". 학년-학기별 가중치 없음. 활용지표는 석차등급.',
    notes:['과목 간/학년 간 세부 가중 방식은 이 자료에서 명확히 확인하지 못했다.'],
    sources:[{...SRC,title:'2025학년도_가천대학교_수시모집요강.pdf',page:84}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2026-humanities']={
    id:'gachon-schoolrecord-2026-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',
    scoreScale:100,
    reflectedSubjects:['국어','수학','영어','사회'],
    yearWeights:{'1학년':1,'2학년':1,'3학년':1},
    gradeToScoreTable:[{grade:1,score:100},{grade:2,score:99.5},{grade:3,score:99.5},{grade:4,score:99},{grade:5,score:99},{grade:6,score:90},{grade:7,score:90},{grade:8,score:70},{grade:9,score:70}],
    method:'인문계열: 국어·수학·영어·사회, 진로선택과목 반영 없음, 3학년 1학기까지 학년·학기 구분 없이 통합 이수단위 가중평균. 석차등급→변환등급(1등급=A, 2·3등급=B, 4·5등급=C, 6·7등급=D, 8·9등급=E) → 배점(A=100, B=99.5, C=99, D=90, E=70).',
    notes:['반영교과·등급→배점 환산표는 모집요강 원문 표(82p)로 완전히 확인했다.','다만 2026학년도 수시 입시결과 파일이 공개하는 "학생부등급(50%/70%/90%)" 컷은 원 석차등급의 단순 평균(1~9 등급 척도)으로 보이며, 위 배점표를 적용한 환산점수(70~100점 척도)와 척도가 달라 서로 비교할 수 없다. 그래서 계산식은 확인됐지만 비교 가능한 공식 컷 기준이 없어 draft로 둔다.']
    ,sources:[{...SRC,title:'2026가천대학교수시모집요강.pdf',page:82}]
  };
  CALCULATION_RULES['gachon-schoolrecord-2026-science']={
    id:'gachon-schoolrecord-2026-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',
    scoreScale:100,
    reflectedSubjects:['국어','수학','영어','과학'],
    yearWeights:{'1학년':1,'2학년':1,'3학년':1},
    gradeToScoreTable:[{grade:1,score:100},{grade:2,score:99.5},{grade:3,score:99},{grade:4,score:98.5},{grade:5,score:98},{grade:6,score:97.5},{grade:7,score:85},{grade:8,score:60},{grade:9,score:30}],
    method:'자연계열: 국어·수학·영어·과학, 진로선택과목 반영 없음, 3학년 1학기까지 학년·학기 구분 없이 통합 이수단위 가중평균. 석차등급→배점 직접 환산(등급별 변환등급 단계 없음): 1=100, 2=99.5, 3=99, 4=98.5, 5=98, 6=97.5, 7=85, 8=60, 9=30.',
    notes:['반영교과·등급→배점 환산표는 모집요강 원문 표(82p)로 완전히 확인했다.','인문계열과 동일한 이유(공개 컷과 배점표 척도 불일치)로 draft로 둔다.']
    ,sources:[{...SRC,title:'2026가천대학교수시모집요강.pdf',page:82}]
  };

  // ===== 모집단위 =====
  const departments=[
    {
      id:'gachon-business',name:'경영학과',college:null,
      simulationBinding:{regular:null,schoolRecord:null,comprehensive:null,essay:null},
      admissionTypes:{
        regular:{details:[
          {id:'gachon-business-regular-general1',name:'일반전형(2026: 일반전형1)',group:'다군',years:[
            {year:2026,calculationRuleId:'gachon-regular-general1-2026',recruitCount:45,competitionRatio:9.44,additionalPass:{count:48,round:null},
              admissionMethod:'수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)',csatMinimum:null,
              indicators:{percentileCut50:87.3,percentileCut70:87.0,percentileCut90:86.8,extra:[]},
              sources:[{...SRC,title:'2026정시입시결과.xls'}]},
            {year:2024,calculationRuleId:'gachon-regular-general1-2024',recruitCount:84,competitionRatio:7.45,additionalPass:{count:207,round:null},
              admissionMethod:'수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)',csatMinimum:null,
              indicators:{percentileCut70:82.30,extra:[{label:'2024학년도 모집단위명',value:'경영학부(2025학년도부터 경영학과로 개편)'}]},
              sources:[{...SRC,title:'2024학년도정시입시결과.xls'}]}
          ]},
          {id:'gachon-business-regular-general2',name:'일반전형2(2026학년도 신설)',group:'다군',years:[
            {year:2026,calculationRuleId:'gachon-regular-general2-2026',recruitCount:17,competitionRatio:32.0,additionalPass:{count:33,round:null},
              admissionMethod:'수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)',csatMinimum:null,
              indicators:{standardScoreCut50:131.4,standardScoreCut70:130.6,standardScoreCut90:130.3,extra:[]},
              sources:[{...SRC,title:'2026정시입시결과.xls'}]}
          ]}
        ]},
        schoolRecord:{details:[
          {id:'gachon-business-schoolrecord-excellent',name:'학생부우수자 전형',group:null,years:[
            {year:2026,calculationRuleId:'gachon-schoolrecord-2026-humanities',recruitCount:15,competitionRatio:14.33,additionalPass:{count:43,round:null},
              admissionMethod:'학생부교과 100%(수능최저학력기준 있음, 세부 기준은 모집요강 참조)',csatMinimum:null,
              indicators:{avgGradeReference50:2.46,avgGradeReference70:2.52,avgGradeReference90:2.63,extra:[{label:'공개 등급 컷의 척도','value':'원 석차등급 평균(1~9등급) — 위 계산규칙의 배점(70~100점) 환산 전 수치로 추정, 직접 비교 불가'}]},
              sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]},
            {year:2025,calculationRuleId:'gachon-schoolrecord-2025-humanities',recruitCount:18,competitionRatio:9.6,additionalPass:{count:43,round:null},
              admissionMethod:'학생부교과 100%',csatMinimum:null,
              indicators:{avgGradeReference70:2.76,avgGradeReference90:2.92,extra:[]},
              sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]},
            {year:2024,calculationRuleId:'gachon-schoolrecord-2024-humanities',recruitCount:36,competitionRatio:10.9,additionalPass:{count:95,round:null},
              admissionMethod:'학생부교과 100%(우수교과 순 40·30·20·10% 반영)',csatMinimum:null,
              indicators:{avgGradeReference70:2.66,avgGradeReference90:2.71,extra:[{label:'2024학년도 모집단위명','value':'경영학부'}]},
              sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]}
          ]}
        ]},
        comprehensive:{details:[]},
        essay:{details:[]}
      }
    },
    {
      id:'gachon-computer',name:'컴퓨터공학과',college:null,
      simulationBinding:{regular:null,schoolRecord:null,comprehensive:null,essay:null},
      admissionTypes:{
        regular:{details:[
          {id:'gachon-computer-regular-general1',name:'일반전형(2026: 일반전형1)',group:'가군',years:[
            {year:2026,calculationRuleId:'gachon-regular-general1-2026',recruitCount:40,competitionRatio:6.2,additionalPass:{count:33,round:null},
              admissionMethod:'수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)',csatMinimum:null,
              indicators:{percentileCut50:87.4,percentileCut70:87.0,percentileCut90:86.8,extra:[]},
              sources:[{...SRC,title:'2026정시입시결과.xls'}]},
            {year:2024,calculationRuleId:'gachon-regular-general1-2024',recruitCount:55,competitionRatio:5.71,additionalPass:{count:105,round:null},
              admissionMethod:'수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)',csatMinimum:null,
              indicators:{percentileCut70:86.46,extra:[{label:'2024학년도 모집단위명','value':'컴퓨터공학전공(2025학년도부터 컴퓨터공학과로 개편)'}]},
              sources:[{...SRC,title:'2024학년도정시입시결과.xls'}]}
          ]},
          {id:'gachon-computer-regular-general2',name:'일반전형2(2026학년도 신설)',group:'가군',years:[
            {year:2026,calculationRuleId:'gachon-regular-general2-2026',recruitCount:18,competitionRatio:13.17,additionalPass:{count:14,round:null},
              admissionMethod:'수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)',csatMinimum:null,
              indicators:{standardScoreCut50:129.2,standardScoreCut70:128.7,standardScoreCut90:128.6,extra:[]},
              sources:[{...SRC,title:'2026정시입시결과.xls'}]}
          ]}
        ]},
        schoolRecord:{details:[
          {id:'gachon-computer-schoolrecord-excellent',name:'학생부우수자 전형',group:null,years:[
            {year:2026,calculationRuleId:'gachon-schoolrecord-2026-science',recruitCount:16,competitionRatio:7.81,additionalPass:{count:19,round:null},
              admissionMethod:'학생부교과 100%(수능최저학력기준 있음, 세부 기준은 모집요강 참조)',csatMinimum:null,
              indicators:{avgGradeReference50:2.6,avgGradeReference70:2.64,avgGradeReference90:2.78,extra:[{label:'공개 등급 컷의 척도','value':'원 석차등급 평균(1~9등급) — 위 계산규칙의 배점(30~100점) 환산 전 수치로 추정, 직접 비교 불가'}]},
              sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]},
            {year:2025,calculationRuleId:'gachon-schoolrecord-2025-science',recruitCount:20,competitionRatio:7.1,additionalPass:{count:28,round:null},
              admissionMethod:'학생부교과 100%',csatMinimum:null,
              indicators:{avgGradeReference70:2.74,avgGradeReference90:2.83,extra:[]},
              sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]},
            {year:2024,calculationRuleId:'gachon-schoolrecord-2024-science',recruitCount:18,competitionRatio:11.9,additionalPass:{count:41,round:null},
              admissionMethod:'학생부교과 100%(우수교과 순 40·30·20·10% 반영)',csatMinimum:null,
              indicators:{avgGradeReference70:2.53,avgGradeReference90:2.58,extra:[{label:'2024학년도 모집단위명','value':'컴퓨터공학전공'}]},
              sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]}
          ]}
        ]},
        comprehensive:{details:[]},
        essay:{details:[]}
      }
    }
  ];

  UNIVERSITY_DB.push({id:'gachon',name:'가천대학교',aliases:['가천대'],region:'경기',departments});
})();
