// data/universities/gachon.js
// 가천대학교 — 대학 구조 DB(UNIVERSITY_DB) 등록 파일. 전체 모집단위/전형 구축분(2단계).
// 이 파일은 build_gachon.js 스크립트로 아래 사용자 제공 원본 파일들을 파싱해 생성했다(수치 임의 변경 없음, 스크립트: 세션 스크래치패드 보관).
//   정시: 2024학년도정시입시결과.xls, 2026정시입시결과.xls, 2026가천대학교정시모집요강수정(단면)_업로드.pdf
//   수시: 가천대학교_2024학년도_수시_입시결과.xlsx, 2025학년도수시입시결과.xlsx, 2026학년도수시입시결과.xlsx, 2026가천대학교수시모집요강.pdf
// 범위: 정시(일반전형1/2) · 학생부교과(학생부우수자·지역균형) · 학생부종합(가천바람개비·가천의약학) · 논술 — 4대 핵심전형만 다루고,
// 실기고사 포함 전형(실기우수자, 정시 연기예술학과 등 예체능계열)과 정원외 특별전형(기회균형·농어촌·특성화고·조기취업 등)은 이번 범위에서 제외했다.
// 2025학년도 정시입시결과.xls는 비밀번호로 보호되어 열지 않았고, 2024학년도 정시모집요강은 .hwp 형식이라 자동 파싱이 불가능해
// 해당 반영방법은 "확인 못함"으로 남겼다(둘 다 임의로 값을 채우지 않음).
(function(){

  // ===== 계산규칙 =====
  const SRC = {org:'가천대학교 입학처(사용자 제공 자료)', url:null, publishedDate:null, checkedDate:'2026-09-01'};
  const ENGLISH_TABLE_2026 = {1:98,2:95,3:92,4:86,5:80,6:60,7:50,8:40,9:30};

  CALCULATION_RULES['gachon-regular-general1-2024']={
    id:'gachon-regular-general1-2024',university:'gachon',admissionType:'regular',appliesToYears:[2024],status:'draft',
    notes:['2024학년도 정시모집요강이 .hwp 형식으로만 제공되어 자동 파싱이 불가능했다. 반영비율을 확인하지 못해 draft로 둔다. 모집인원·경쟁률·70%컷은 원자료 그대로 등록했다.'],
    sources:[{...SRC,title:'2024학년도정시모집요강.hwp(파싱 불가)'},{...SRC,title:'2024학년도정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-regular-general1-2026']={
    id:'gachon-regular-general1-2026',university:'gachon',admissionType:'regular',appliesToYears:[2026],status:'verified',
    engine:'ranked-subject-weighted',
    rankedSubjects:['kor','mat','eng'],rankedWeights:[35,25,20],fixedSubjects:{inq:20},
    englishGradeToScore:ENGLISH_TABLE_2026,
    method:'인문·자연계열(의예과·약학과·한의예과·클라우드공학과 제외) 공통: 국어·수학·영어 중 백분위(영어는 등급환산점수)가 높은 순서로 35%·25%·20% 반영 + 탐구(1과목) 백분위 20% 반영. 한국사는 응시 필수(가산/감점 없음, 원문에 별도 배점표 없음).',
    notes:['영어 등급→배점표(1~9등급: 98,95,92,86,80,60,50,40,30)는 모집요강 원문(43p)으로 완전히 확인했다.','클라우드공학과는 이 표에서 제외 대상으로 명시되어 있으나 별도 반영비율을 확인하지 못해 클라우드공학과 정시는 draft로 둔다(계산 보류).','한국사는 "필수"로만 표기되고 등급별 가산·감점표가 원문에 없어 응시 요건으로만 반영(성적 계산에는 영향 없음)으로 해석했다.'],
    sources:[{...SRC,title:'2026가천대학교정시모집요강수정(단면)_업로드.pdf',page:41},{...SRC,title:'2026정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-regular-general2-2026']={
    id:'gachon-regular-general2-2026',university:'gachon',admissionType:'regular',appliesToYears:[2026],status:'draft',
    reflectedSubjects:['국어','수학','영어','탐구'],
    method:'2026학년도 신설. 국어·수학·영어·탐구(1과목) 중 우수한 영역순 50%·30%·20%만 반영(1개 영역 미반영). 활용지표는 표준점수(영어는 변환표준점수, 국어·수학 등급구간 표준점수 평균으로 산출). 탐구는 2과목 중 우수한 과목의 표준점수×2.',
    notes:['표준점수는 그 해 응시생 집단에 따라 달라지는 상대값이라 사전에 고정된 환산표가 존재하지 않는다(대학도 "수능 성적 발표 후 입학처 홈페이지에서 확인"이라고 원문에 명시). 구조적으로 verified 처리가 불가능해 draft로 둔다.'],
    sources:[{...SRC,title:'2026가천대학교정시모집요강수정(단면)_업로드.pdf',page:41},{...SRC,title:'2026정시입시결과.xls'}]
  };
  CALCULATION_RULES['gachon-regular-fixed-2026']={
    id:'gachon-regular-fixed-2026',university:'gachon',admissionType:'regular',appliesToYears:[2026],status:'verified',
    engine:'ranked-subject-weighted',
    rankedSubjects:[],rankedWeights:[],fixedSubjects:{kor:25,mat:30,eng:20,inq:25},
    englishGradeToScore:ENGLISH_TABLE_2026,
    method:'의예과·약학과·한의예과: 국어25%+수학30%+영어20%+탐구2과목25%(과학탐구, 우수영역순 아님, 고정비율). 한국사 응시 필수(가산/감점 없음).',
    notes:['탐구는 앱의 기존 "탐구 평균 백분위" 입력값(2과목 평균)을 그대로 사용한다(원문도 2과목 평균 방식).'],
    sources:[{...SRC,title:'2026가천대학교정시모집요강수정(단면)_업로드.pdf',page:41},{...SRC,title:'2026정시입시결과.xls'}]
  };

  // ===== 학생부교과(학생부우수자·지역균형) 계산규칙 — 전부 draft(계산 보류), 이유는 각 notes 참고 =====
  CALCULATION_RULES['gachon-schoolrecord-excellent-2024-humanities']={id:'gachon-schoolrecord-excellent-2024-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"학생부우수자(2024): 국어·수학·영어·사회 중 우수한 교과 순으로 40%·30%·20%·10% 반영(진로선택 미반영). 등급→점수 환산표를 확인하지 못해 draft로 둔다.",sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-excellent-2024-science']={id:'gachon-schoolrecord-excellent-2024-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"학생부우수자(2024): 국어·수학·영어·과학 중 우수한 교과 순으로 40%·30%·20%·10% 반영(진로선택 미반영). 등급→점수 환산표를 확인하지 못해 draft로 둔다.",sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2024-humanities']={id:'gachon-schoolrecord-balanced-2024-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2024-science']={id:'gachon-schoolrecord-balanced-2024-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2024],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'가천대학교_2024학년도_수시_입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-excellent-2025-humanities']={id:'gachon-schoolrecord-excellent-2025-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"학생부우수자(2025): 국어·수학·영어·사회, 교과 100% 반영(요소기준), 학년-학기 가중치 없음. 세부 등급→점수 환산 방식을 확인하지 못해 draft로 둔다.",sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-excellent-2025-science']={id:'gachon-schoolrecord-excellent-2025-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"학생부우수자(2025): 국어·수학·영어·과학, 교과 100% 반영(요소기준), 학년-학기 가중치 없음. 세부 등급→점수 환산 방식을 확인하지 못해 draft로 둔다.",sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2025-humanities']={id:'gachon-schoolrecord-balanced-2025-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2025-science']={id:'gachon-schoolrecord-balanced-2025-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2025],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'2025학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-excellent-2026-humanities']={id:'gachon-schoolrecord-excellent-2026-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"학생부우수자(2026): 국어·수학·영어·사회, 3학년 1학기까지 학년 구분 없이 통합 이수단위 가중평균. 등급→배점 환산표는 모집요강(82p)으로 확인했지만, 공개된 결과 컷(석차등급 평균 척도)과 실제 배점(70~100/30~100점 척도)의 단위가 달라 비교가 불가능해 draft로 둔다.",sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-excellent-2026-science']={id:'gachon-schoolrecord-excellent-2026-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"학생부우수자(2026): 국어·수학·영어·과학, 3학년 1학기까지 학년 구분 없이 통합 이수단위 가중평균. 등급→배점 환산표는 모집요강(82p)으로 확인했지만, 공개된 결과 컷(석차등급 평균 척도)과 실제 배점(70~100/30~100점 척도)의 단위가 달라 비교가 불가능해 draft로 둔다.",sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2026-humanities']={id:'gachon-schoolrecord-balanced-2026-humanities',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',reflectedSubjects:['국어','수학','영어','사회'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]};
  CALCULATION_RULES['gachon-schoolrecord-balanced-2026-science']={id:'gachon-schoolrecord-balanced-2026-science',university:'gachon',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',reflectedSubjects:['국어','수학','영어','과학'],method:"지역균형: 1단계 학생부교과 100%(7배수, 진로선택과목 성취도만 반영) → 2단계 1단계평가 50% + 면접 50%. 2단계 면접이 정성평가라 정량 재현이 불가능해 draft로 둔다.",sources:[{...SRC,title:'2026학년도수시입시결과.xlsx'}]};

  const departments = [
   {
    "id": "gachon-01",
    "name": "간호학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-01-regular-general1",
     "schoolRecord": "gachon-01-schoolrecord-excellent",
     "comprehensive": "gachon-01-comprehensive-baram",
     "essay": "gachon-01-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-01-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 71,
          "competitionRatio": 3.59,
          "additionalPass": {
           "count": 30,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 89.8,
           "percentileCut70": 89.5,
           "percentileCut90": 89.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 61,
          "competitionRatio": 4.25,
          "additionalPass": {
           "count": 60,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 88.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-01-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 30,
          "competitionRatio": 6.77,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 131.7,
           "standardScoreCut70": 131.3,
           "standardScoreCut90": 131.1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-01-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 28,
          "competitionRatio": 7.5,
          "additionalPass": {
           "count": 38,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 1.89,
           "avgGradeReference70": 1.99,
           "avgGradeReference90": 2.13,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 33,
          "competitionRatio": 6.9,
          "additionalPass": {
           "count": 53,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.01,
           "avgGradeReference90": 2.07,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 42,
          "competitionRatio": 7.38,
          "additionalPass": {
           "count": 84,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.9,
           "avgGradeReference90": 2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-01-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-science",
          "recruitCount": 13,
          "competitionRatio": 40,
          "additionalPass": {
           "count": 3.33,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 2.3,
           "avgGradeReference70": 2.32,
           "avgGradeReference90": 2.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 15,
          "competitionRatio": 47.1,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.1,
           "avgGradeReference90": 4.39,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 20,
          "competitionRatio": 13.6,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.22,
           "avgGradeReference90": 2.32,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-01-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 33,
          "competitionRatio": 28.18,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 2.79,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 3.28,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 38,
          "competitionRatio": 23.8,
          "additionalPass": {
           "count": 24,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.8,
           "avgGradeReference90": 3.23,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 45,
          "competitionRatio": 26.53,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.82,
           "avgGradeReference90": 2.99,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-01-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 77,
          "competitionRatio": 47.68,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.49,
           "avgGradeReference90": 5.6,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=118.5, 90%=117"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 83,
          "competitionRatio": 44.6,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.26,
           "avgGradeReference90": 5.01,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=111, 90%=109"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 83,
          "competitionRatio": 46.36,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.88,
           "avgGradeReference90": 4.6,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.9, 90%=12.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-02",
    "name": "건설환경공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-02-regular-general1",
     "schoolRecord": "gachon-02-schoolrecord-excellent",
     "comprehensive": "gachon-02-comprehensive-baram",
     "essay": "gachon-02-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-02-regular-general1",
        "name": "일반전형1",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 17,
          "competitionRatio": 14.29,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86,
           "percentileCut70": 85.8,
           "percentileCut90": 85.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-02-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 10.57,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.81,
           "avgGradeReference70": 2.86,
           "avgGradeReference90": 2.97,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-02-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 24.67,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.63,
           "avgGradeReference70": 3.78,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-02-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 26.43,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.43,
           "avgGradeReference90": 5.79,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-03",
    "name": "건축공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-03-regular-general1",
     "schoolRecord": "gachon-03-schoolrecord-excellent",
     "comprehensive": "gachon-03-comprehensive-baram",
     "essay": "gachon-03-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-03-regular-general1",
        "name": "일반전형1",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 12,
          "competitionRatio": 11.25,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.9,
           "percentileCut70": 86.8,
           "percentileCut90": 86.7,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-03-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 14.17,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 3.07,
           "avgGradeReference70": 3.17,
           "avgGradeReference90": 3.21,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-03-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 20.14,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.61,
           "avgGradeReference70": 3.92,
           "avgGradeReference90": 3.95,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-03-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 28.07,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.07,
           "avgGradeReference90": 5.76,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=109, 90%=105"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-04",
    "name": "건축학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-04-regular-general1",
     "schoolRecord": "gachon-04-schoolrecord-excellent",
     "comprehensive": "gachon-04-comprehensive-baram",
     "essay": "gachon-04-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-04-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 18,
          "competitionRatio": 8.61,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.4,
           "percentileCut70": 87.1,
           "percentileCut90": 87,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 52,
          "competitionRatio": 7.34,
          "additionalPass": {
           "count": 117,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.87,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-04-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 11,
          "competitionRatio": 21.55,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 130,
           "standardScoreCut70": 129.5,
           "standardScoreCut90": 129.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-04-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 23.71,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.72,
           "avgGradeReference70": 2.81,
           "avgGradeReference90": 2.94,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 13,
          "competitionRatio": 8.1,
          "additionalPass": {
           "count": 30,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.3,
           "avgGradeReference90": 3.54,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 14,
          "competitionRatio": 13.86,
          "additionalPass": {
           "count": 26,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.77,
           "avgGradeReference90": 2.79,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-04-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 27,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.47,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 10,
          "competitionRatio": 18.3,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.42,
           "avgGradeReference90": 3.45,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-04-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 39.5,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.55,
           "avgGradeReference70": 3.94,
           "avgGradeReference90": 4.41,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 37.6,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.94,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 37,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.83,
           "avgGradeReference90": 3.86,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-04-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 52.4,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.46,
           "avgGradeReference90": 5.98,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=114, 90%=111"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 33,
          "competitionRatio": 37.4,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.16,
           "avgGradeReference90": 5.93,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103.5, 90%=102"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 30,
          "competitionRatio": 37,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.74,
           "avgGradeReference90": 5.17,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.5, 90%=11.2"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-05",
    "name": "경영학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-05-regular-general1",
     "schoolRecord": "gachon-05-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-05-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-05-regular-general1",
        "name": "일반전형1",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 45,
          "competitionRatio": 9.44,
          "additionalPass": {
           "count": 48,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.3,
           "percentileCut70": 87,
           "percentileCut90": 86.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-05-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 17,
          "competitionRatio": 32,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 131.4,
           "standardScoreCut70": 130.6,
           "standardScoreCut90": 130.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-05-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 15,
          "competitionRatio": 14.33,
          "additionalPass": {
           "count": 43,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.46,
           "avgGradeReference70": 2.52,
           "avgGradeReference90": 2.63,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 18,
          "competitionRatio": 9.6,
          "additionalPass": {
           "count": 43,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.76,
           "avgGradeReference90": 2.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 36,
          "competitionRatio": 10.89,
          "additionalPass": {
           "count": 95,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.66,
           "avgGradeReference90": 2.71,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "경영학부"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-05-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 18,
          "competitionRatio": 11.56,
          "additionalPass": {
           "count": 31,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.18,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "경영학부"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": []
     },
     "essay": {
      "details": [
       {
        "id": "gachon-05-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 45,
          "competitionRatio": 58.73,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.23,
           "avgGradeReference90": 5.98,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=113, 90%=111"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 50,
          "competitionRatio": 47.2,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.37,
           "avgGradeReference90": 6.21,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=114, 90%=112"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 78,
          "competitionRatio": 31.9,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.65,
           "avgGradeReference90": 5.33,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=13, 90%=12.7"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "경영학부"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-06",
    "name": "경제학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-06-schoolrecord-excellent",
     "comprehensive": "gachon-06-comprehensive-baram",
     "essay": "gachon-06-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": []
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-06-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 8,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.76,
           "avgGradeReference70": 2.8,
           "avgGradeReference90": 2.82,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 7,
          "competitionRatio": 9.4,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.8,
           "avgGradeReference90": 3.19,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.65,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-06-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 1,
          "competitionRatio": 19,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 6.81,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15.33,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.14,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-06-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 22.3,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.06,
           "avgGradeReference90": 4.38,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 27.13,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.03,
           "avgGradeReference90": 4.22,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-06-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 43.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.22,
           "avgGradeReference90": 6.49,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110, 90%=108"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 35.1,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.93,
           "avgGradeReference90": 5.36,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110, 90%=108.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 35.36,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.41,
           "avgGradeReference90": 5.26,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.8, 90%=12.8"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-07",
    "name": "경찰행정학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-07-schoolrecord-excellent",
     "comprehensive": "gachon-07-comprehensive-baram",
     "essay": "gachon-07-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-07-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 7.56,
          "additionalPass": {
           "count": 25,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.7,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)경찰행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(경찰행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-07-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 11.17,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.19,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)경찰행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(경찰행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-07-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 10.17,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.74,
           "avgGradeReference90": 2.89,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)경찰행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(경찰행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-07-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 28.67,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.23,
           "avgGradeReference90": 3.74,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)경찰행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(경찰행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-07-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 34.1,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.45,
           "avgGradeReference90": 3.79,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.1, 90%=12.1"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)경찰행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(경찰행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-08",
    "name": "관광경영학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-08-schoolrecord-excellent",
     "comprehensive": "gachon-08-comprehensive-baram",
     "essay": "gachon-08-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": []
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-08-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 7,
          "competitionRatio": 11.86,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.94,
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.07,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 7,
          "competitionRatio": 9.4,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3,
           "avgGradeReference90": 3.17,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15.83,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.69,
           "avgGradeReference90": 2.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-08-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 2,
          "competitionRatio": 29,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.49,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 10,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.37,
           "avgGradeReference90": 3.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-08-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 60.13,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.29,
           "avgGradeReference90": 4.62,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-08-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 43.38,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.49,
           "avgGradeReference90": 5.57,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=104"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 36.6,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.07,
           "avgGradeReference90": 5.94,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=114.5, 90%=112.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 32.73,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.21,
           "avgGradeReference90": 4.41,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.3, 90%=11"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-09",
    "name": "금융·빅데이터학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-09-schoolrecord-excellent",
     "comprehensive": "gachon-09-comprehensive-baram",
     "essay": "gachon-09-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-09-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 40,
          "competitionRatio": 7.45,
          "additionalPass": {
           "count": 113,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-09-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 12,
          "competitionRatio": 20.5,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.3,
           "standardScoreCut70": 129.2,
           "standardScoreCut90": 128.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-09-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 12,
          "competitionRatio": 12.5,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.88,
           "avgGradeReference70": 2.96,
           "avgGradeReference90": 3.09,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 16,
          "competitionRatio": 7,
          "additionalPass": {
           "count": 23,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.17,
           "avgGradeReference90": 3.36,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-09-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 15.9,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.02,
           "avgGradeReference90": 4.77,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-09-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 26,
          "competitionRatio": 29.08,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.63,
           "avgGradeReference90": 5.35,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110, 90%=108"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 26,
          "competitionRatio": 27.5,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.2,
           "avgGradeReference90": 5.98,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107.5, 90%=105"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-10",
    "name": "금융수학전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-10-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-10-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": []
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-10-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 9.17,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 2.97,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)금융수학전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(금융수학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-10-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 9.83,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.19,
           "avgGradeReference90": 3.63,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)금융수학전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(금융수학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": []
     },
     "essay": {
      "details": [
       {
        "id": "gachon-10-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 25.9,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4,
           "avgGradeReference90": 4.72,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.5, 90%=11.4"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)금융수학전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(금융수학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-11",
    "name": "기계공학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-11-regular-general1",
     "schoolRecord": "gachon-11-schoolrecord-excellent",
     "comprehensive": "gachon-11-comprehensive-baram",
     "essay": "gachon-11-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-11-regular-general1",
        "name": "일반전형1",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 45,
          "competitionRatio": 8.44,
          "additionalPass": {
           "count": 56,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.2,
           "percentileCut70": 87,
           "percentileCut90": 86.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-11-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 22,
          "competitionRatio": 19.05,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.3,
           "standardScoreCut70": 129.3,
           "standardScoreCut90": 129,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-11-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 18,
          "competitionRatio": 11.94,
          "additionalPass": {
           "count": 34,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.73,
           "avgGradeReference70": 2.79,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 26,
          "competitionRatio": 10.7,
          "additionalPass": {
           "count": 31,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.93,
           "avgGradeReference90": 3.13,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-11-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 18,
          "competitionRatio": 35.72,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.62,
           "avgGradeReference70": 3.95,
           "avgGradeReference90": 4.11,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 28,
          "competitionRatio": 22.3,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.11,
           "avgGradeReference90": 4.46,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-11-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 56,
          "competitionRatio": 32.45,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.12,
           "avgGradeReference90": 5.8,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=112, 90%=109.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 51,
          "competitionRatio": 32,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.18,
           "avgGradeReference90": 6.18,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103, 90%=101.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-12",
    "name": "기계공학전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-12-schoolrecord-excellent",
     "comprehensive": "gachon-12-comprehensive-baram",
     "essay": "gachon-12-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-12-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 41,
          "competitionRatio": 6.78,
          "additionalPass": {
           "count": 129,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.54,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)기계공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(기계공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-12-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 12,
          "competitionRatio": 45.25,
          "additionalPass": {
           "count": 39,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.89,
           "avgGradeReference90": 3,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)기계공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(기계공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-12-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 8,
          "competitionRatio": 12.63,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.18,
           "avgGradeReference90": 3.19,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)기계공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(기계공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-12-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 26.44,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.16,
           "avgGradeReference90": 4.76,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)기계공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(기계공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-12-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 35.52,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.33,
           "avgGradeReference90": 4.53,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.7, 90%=11.5"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)기계공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(기계공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-13",
    "name": "도시계획·조경학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-13-schoolrecord-excellent",
     "comprehensive": "gachon-13-comprehensive-baram",
     "essay": "gachon-13-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-13-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 33,
          "competitionRatio": 7.58,
          "additionalPass": {
           "count": 84,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.77,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-13-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 10,
          "competitionRatio": 19.1,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.91,
           "avgGradeReference70": 2.95,
           "avgGradeReference90": 3.02,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 10,
          "competitionRatio": 9,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.51,
           "avgGradeReference90": 3.56,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 10,
          "competitionRatio": 14.1,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.76,
           "avgGradeReference90": 2.94,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-13-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.73,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 7,
          "competitionRatio": 17.29,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.44,
           "avgGradeReference90": 3.63,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-13-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 30.86,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.03,
           "avgGradeReference70": 4.15,
           "avgGradeReference90": 4.39,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 25.9,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.22,
           "avgGradeReference90": 4.28,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 21.29,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.87,
           "avgGradeReference90": 4.26,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-13-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 28.43,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.99,
           "avgGradeReference90": 6.23,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104.5, 90%=103"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 30.7,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.01,
           "avgGradeReference90": 5.51,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=93, 90%=91.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 20,
          "competitionRatio": 27.75,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.03,
           "avgGradeReference90": 5.31,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.1, 90%=10.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-14",
    "name": "물리치료학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-14-regular-general1",
     "schoolRecord": "gachon-14-schoolrecord-excellent",
     "comprehensive": "gachon-14-comprehensive-baram",
     "essay": "gachon-14-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-14-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 14,
          "competitionRatio": 8.71,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 88.5,
           "percentileCut70": 88.4,
           "percentileCut90": 87.5,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 14,
          "competitionRatio": 8.5,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 86.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-14-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 16.83,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 1.98,
           "avgGradeReference70": 2.3,
           "avgGradeReference90": 2.41,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 6,
          "competitionRatio": 10.3,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.11,
           "avgGradeReference90": 2.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.4,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.79,
           "avgGradeReference90": 2.01,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-14-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-science",
          "recruitCount": 6,
          "competitionRatio": 21,
          "additionalPass": {
           "count": 3.05,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 2.17,
           "avgGradeReference70": 2.41,
           "avgGradeReference90": 2.7,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 6,
          "competitionRatio": 38.5,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.65,
           "avgGradeReference90": 3.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.8,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.25,
           "avgGradeReference90": 2.27,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-14-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 62.67,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.31,
           "avgGradeReference70": 3.48,
           "avgGradeReference90": 3.48,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 37.3,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.95,
           "avgGradeReference90": 4.22,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 57.17,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.9,
           "avgGradeReference90": 2.96,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-14-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 76.25,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.1,
           "avgGradeReference90": 4.67,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=115, 90%=114"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 78.8,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.75,
           "avgGradeReference90": 5.2,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110.75, 90%=110"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 70.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.78,
           "avgGradeReference90": 4.62,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.2, 90%=10.1"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-15",
    "name": "미디어커뮤니케이션학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-15-regular-general1",
     "schoolRecord": "gachon-15-schoolrecord-excellent",
     "comprehensive": "gachon-15-comprehensive-baram",
     "essay": "gachon-15-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-15-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 11,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.6,
           "percentileCut70": 87.1,
           "percentileCut90": 86.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 21,
          "competitionRatio": 8.19,
          "additionalPass": {
           "count": 65,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.55,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-15-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 5,
          "competitionRatio": 13.6,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2,
           "avgGradeReference70": 2.66,
           "avgGradeReference90": 3.19,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 5,
          "competitionRatio": 11.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.32,
           "avgGradeReference90": 2.75,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 7,
          "competitionRatio": 21.14,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.44,
           "avgGradeReference90": 2.46,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-15-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 1,
          "competitionRatio": 31,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 6.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 18.5,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.42,
           "avgGradeReference90": 2.94,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-15-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 114.67,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.54,
           "avgGradeReference70": 3.86,
           "avgGradeReference90": 3.89,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 62,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.97,
           "avgGradeReference90": 4.76,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 64.64,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.69,
           "avgGradeReference90": 3.8,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-15-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 83.75,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.97,
           "avgGradeReference90": 5.53,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=117, 90%=116"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 62.4,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.33,
           "avgGradeReference90": 6.87,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=118, 90%=117"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 52.92,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.29,
           "avgGradeReference90": 5.11,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=13.2, 90%=12.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-16",
    "name": "바이오나노학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-16-regular-general1",
     "schoolRecord": "gachon-16-schoolrecord-excellent",
     "comprehensive": "gachon-16-comprehensive-baram",
     "essay": "gachon-16-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-16-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 16,
          "competitionRatio": 7.13,
          "additionalPass": {
           "count": 28,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.6,
           "percentileCut70": 86.4,
           "percentileCut90": 86.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 7.38,
          "additionalPass": {
           "count": 44,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.75,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-16-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 18.14,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.68,
           "avgGradeReference70": 2.81,
           "avgGradeReference90": 2.88,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 5.9,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.16,
           "avgGradeReference90": 3.38,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 7.57,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.38,
           "avgGradeReference90": 2.61,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-16-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 11.83,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.81,
           "avgGradeReference90": 3.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-16-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 24,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.01,
           "avgGradeReference70": 3.15,
           "avgGradeReference90": 3.26,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 25.7,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.4,
           "avgGradeReference90": 3.43,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 32,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.42,
           "avgGradeReference90": 3.73,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-16-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 29.5,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.58,
           "avgGradeReference90": 5.21,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=113, 90%=111"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 32.1,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.24,
           "avgGradeReference90": 4.62,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=105, 90%=103"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 30.85,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.7,
           "avgGradeReference90": 5.73,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.4, 90%=11.1"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-17",
    "name": "바이오로직스학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-17-regular-general1",
     "schoolRecord": "gachon-17-schoolrecord-excellent",
     "comprehensive": "gachon-17-comprehensive-baram",
     "essay": "gachon-17-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-17-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 30,
          "competitionRatio": 9.2,
          "additionalPass": {
           "count": 39,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.5,
           "percentileCut70": 87.4,
           "percentileCut90": 86.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 40,
          "competitionRatio": 5.56,
          "additionalPass": {
           "count": 63,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 86.12,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-17-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 15,
          "competitionRatio": 18.33,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 130.2,
           "standardScoreCut70": 129.8,
           "standardScoreCut90": 129.5,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-17-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 12,
          "competitionRatio": 12.5,
          "additionalPass": {
           "count": 22,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.66,
           "avgGradeReference70": 2.68,
           "avgGradeReference90": 2.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 16,
          "competitionRatio": 6.4,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.9,
           "avgGradeReference90": 3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 20,
          "competitionRatio": 9.15,
          "additionalPass": {
           "count": 30,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.41,
           "avgGradeReference90": 2.64,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-17-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 11.8,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.19,
           "avgGradeReference70": 3.33,
           "avgGradeReference90": 3.69,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 9.6,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.06,
           "avgGradeReference90": 3.3,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 18.67,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.75,
           "avgGradeReference90": 3.27,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-17-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 28,
          "competitionRatio": 33.36,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.86,
           "avgGradeReference90": 6.34,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110, 90%=108"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 28,
          "competitionRatio": 30.6,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.75,
           "avgGradeReference90": 5.3,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=105, 90%=104"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 25,
          "competitionRatio": 27.24,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.44,
           "avgGradeReference90": 5.45,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.2, 90%=9.8"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-18",
    "name": "반도체·전자공학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-18-schoolrecord-excellent",
     "comprehensive": "gachon-18-comprehensive-baram",
     "essay": "gachon-18-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-18-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 101,
          "competitionRatio": 5.4,
          "additionalPass": {
           "count": 139,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.85,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(반도체대학,시스템반도체학과)반도체·전자공학부\" — 여러 단위로 분리·개편된 것으로 보이나 어느 쪽으로 이어지는지 특정할 수 없어 임의 연결하지 않음"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-18-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 26,
          "competitionRatio": 11.54,
          "additionalPass": {
           "count": 77,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.7,
           "avgGradeReference90": 2.76,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(반도체대학,시스템반도체학과)반도체·전자공학부\" — 여러 단위로 분리·개편된 것으로 보이나 어느 쪽으로 이어지는지 특정할 수 없어 임의 연결하지 않음"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-18-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 16,
          "competitionRatio": 11.31,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.22,
           "avgGradeReference90": 3.3,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(반도체대학,시스템반도체학과)반도체·전자공학부\" — 여러 단위로 분리·개편된 것으로 보이나 어느 쪽으로 이어지는지 특정할 수 없어 임의 연결하지 않음"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-18-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 31,
          "competitionRatio": 14.23,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.69,
           "avgGradeReference90": 4.06,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(반도체대학,시스템반도체학과)반도체·전자공학부\" — 여러 단위로 분리·개편된 것으로 보이나 어느 쪽으로 이어지는지 특정할 수 없어 임의 연결하지 않음"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-18-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 51,
          "competitionRatio": 39.37,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.63,
           "avgGradeReference90": 5.27,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.3, 90%=10"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(반도체대학,시스템반도체학과)반도체·전자공학부\" — 여러 단위로 분리·개편된 것으로 보이나 어느 쪽으로 이어지는지 특정할 수 없어 임의 연결하지 않음"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-19",
    "name": "반도체대학",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-19-regular-general1",
     "schoolRecord": "gachon-19-schoolrecord-excellent",
     "comprehensive": "gachon-19-comprehensive-baram",
     "essay": "gachon-19-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-19-regular-general1",
        "name": "일반전형1",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 61,
          "competitionRatio": 5.1,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 88.5,
           "percentileCut70": 88.2,
           "percentileCut90": 87.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-19-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 24,
          "competitionRatio": 14.63,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 131.3,
           "standardScoreCut70": 130.9,
           "standardScoreCut90": 130.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-19-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 21,
          "competitionRatio": 11.14,
          "additionalPass": {
           "count": 43,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.64,
           "avgGradeReference70": 2.82,
           "avgGradeReference90": 2.89,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 30,
          "competitionRatio": 6.2,
          "additionalPass": {
           "count": 34,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.24,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-19-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 17,
          "competitionRatio": 16.12,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.39,
           "avgGradeReference70": 3.55,
           "avgGradeReference90": 3.65,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 32,
          "competitionRatio": 9.6,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.7,
           "avgGradeReference90": 3.95,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-19-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 73,
          "competitionRatio": 33.37,
          "additionalPass": {
           "count": 18,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.22,
           "avgGradeReference90": 6.16,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=109, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 61,
          "competitionRatio": 34.5,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.93,
           "avgGradeReference90": 5.42,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=95.5, 90%=93.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-20",
    "name": "반도체물리학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-20-schoolrecord-excellent",
     "comprehensive": "gachon-20-comprehensive-baram",
     "essay": "gachon-20-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-20-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 6.53,
          "additionalPass": {
           "count": 47,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.12,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "물리학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-20-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.76,
           "avgGradeReference70": 2.92,
           "avgGradeReference90": 2.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 7.3,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.02,
           "avgGradeReference90": 3.16,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 13.67,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.87,
           "avgGradeReference90": 3.24,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "물리학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-20-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 2,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.49,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 11.33,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.69,
           "avgGradeReference90": 3.79,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "물리학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-20-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 13.69,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.31,
           "avgGradeReference70": 4.34,
           "avgGradeReference90": 4.88,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 13.1,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.84,
           "avgGradeReference90": 4.32,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 14.14,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.18,
           "avgGradeReference90": 4.23,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "물리학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-20-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 25.77,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.64,
           "avgGradeReference90": 5.76,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104, 90%=102"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 30.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.47,
           "avgGradeReference90": 5.73,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=102, 90%=98"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 27,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.43,
           "avgGradeReference90": 5.55,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.6, 90%=10.6"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "물리학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-21",
    "name": "방사선학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-21-regular-general1",
     "schoolRecord": "gachon-21-schoolrecord-excellent",
     "comprehensive": "gachon-21-comprehensive-baram",
     "essay": "gachon-21-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-21-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 14,
          "competitionRatio": 9.07,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 89.4,
           "percentileCut70": 89.2,
           "percentileCut90": 88.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 14,
          "competitionRatio": 7.14,
          "additionalPass": {
           "count": 24,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 87.15,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-21-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 11.17,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 1.78,
           "avgGradeReference70": 1.96,
           "avgGradeReference90": 2.06,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 6,
          "competitionRatio": 10,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.85,
           "avgGradeReference90": 2.15,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 20,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2,
           "avgGradeReference90": 2.03,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-21-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-science",
          "recruitCount": 6,
          "competitionRatio": 23.17,
          "additionalPass": {
           "count": 3.53,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 2.52,
           "avgGradeReference70": 2.65,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 6,
          "competitionRatio": 25,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.08,
           "avgGradeReference90": 4.13,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 20.4,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.78,
           "avgGradeReference90": 2.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-21-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 42,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.05,
           "avgGradeReference70": 3.05,
           "avgGradeReference90": 3.09,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 38.2,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.8,
           "avgGradeReference90": 3.86,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 38.67,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.78,
           "avgGradeReference90": 3.94,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-21-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 64.25,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 6.07,
           "avgGradeReference90": 6.17,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=123, 90%=117"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 59,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.62,
           "avgGradeReference90": 4.99,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110, 90%=110"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 42.7,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.58,
           "avgGradeReference90": 5.08,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.1, 90%=11.9"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-22",
    "name": "법과대학",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-22-schoolrecord-excellent",
     "comprehensive": "gachon-22-comprehensive-baram",
     "essay": "gachon-22-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-22-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 18,
          "competitionRatio": 16.17,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.9,
           "standardScoreCut70": 129.7,
           "standardScoreCut90": 129.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-22-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 17,
          "competitionRatio": 12.82,
          "additionalPass": {
           "count": 40,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.67,
           "avgGradeReference70": 2.8,
           "avgGradeReference90": 2.86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 22,
          "competitionRatio": 10.5,
          "additionalPass": {
           "count": 53,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.84,
           "avgGradeReference90": 2.97,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-22-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 22,
          "competitionRatio": 36.05,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.56,
           "avgGradeReference70": 3.64,
           "avgGradeReference90": 3.85,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 26,
          "competitionRatio": 25.5,
          "additionalPass": {
           "count": 27,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.8,
           "avgGradeReference90": 4.32,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-22-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 45,
          "competitionRatio": 42.16,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.85,
           "avgGradeReference90": 5.5,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=117, 90%=116"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 44,
          "competitionRatio": 32.8,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.67,
           "avgGradeReference90": 5.66,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=109.5, 90%=108"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-23",
    "name": "법학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-23-schoolrecord-excellent",
     "comprehensive": "gachon-23-comprehensive-baram",
     "essay": "gachon-23-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-23-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 46,
          "competitionRatio": 6.39,
          "additionalPass": {
           "count": 113,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.2,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)법학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(법학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-23-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 9,
          "competitionRatio": 12.11,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.73,
           "avgGradeReference90": 2.82,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)법학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(법학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-23-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 9,
          "competitionRatio": 12.67,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.13,
           "avgGradeReference90": 3.27,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)법학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(법학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-23-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 28,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.83,
           "avgGradeReference90": 4.19,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)법학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(법학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-23-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 22,
          "competitionRatio": 29.5,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.91,
           "avgGradeReference90": 5.34,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.9, 90%=12.5"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)법학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(법학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-24",
    "name": "빅데이터경영전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-24-schoolrecord-excellent",
     "comprehensive": "gachon-24-comprehensive-baram",
     "essay": "gachon-24-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": []
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-24-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 13,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.1,
           "avgGradeReference90": 3.17,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)빅데이터경영전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(빅데이터경영전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-24-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 9.33,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.33,
           "avgGradeReference90": 3.75,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)빅데이터경영전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(빅데이터경영전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-24-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 16.57,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.43,
           "avgGradeReference90": 5.24,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)빅데이터경영전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(빅데이터경영전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-24-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 28.4,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.66,
           "avgGradeReference90": 4.55,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.2, 90%=10.6"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(금융·빅데이터학부)빅데이터경영전공\" — 공식 자료상 금융·빅데이터학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 금융·빅데이터학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(빅데이터경영전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-25",
    "name": "사회복지학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-25-regular-general1",
     "schoolRecord": "gachon-25-schoolrecord-excellent",
     "comprehensive": "gachon-25-comprehensive-baram",
     "essay": "gachon-25-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-25-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 13,
          "competitionRatio": 8.77,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 85.5,
           "percentileCut70": 85.5,
           "percentileCut90": 85.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 15,
          "competitionRatio": 9.56,
          "additionalPass": {
           "count": 45,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-25-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 6,
          "competitionRatio": 25.5,
          "additionalPass": {
           "count": 18,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.88,
           "avgGradeReference70": 2.94,
           "avgGradeReference90": 2.94,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 17.7,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.17,
           "avgGradeReference90": 3.22,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 36.5,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3,
           "avgGradeReference90": 3.11,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-25-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 1,
          "competitionRatio": 20,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.47,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 11.33,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.23,
           "avgGradeReference90": 3.29,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-25-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 43.14,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.06,
           "avgGradeReference70": 4.14,
           "avgGradeReference90": 4.16,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 47.6,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.71,
           "avgGradeReference90": 4.01,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 33.56,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.2,
           "avgGradeReference90": 4.3,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-25-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 40.33,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.63,
           "avgGradeReference90": 6.49,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=105.5, 90%=104"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 34,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5,
           "avgGradeReference90": 5.98,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=103"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 29.18,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.13,
           "avgGradeReference90": 4.73,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.6, 90%=12.6"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-26",
    "name": "산업공학전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-26-schoolrecord-excellent",
     "comprehensive": "gachon-26-comprehensive-baram",
     "essay": "gachon-26-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-26-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 18,
          "competitionRatio": 7,
          "additionalPass": {
           "count": 29,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.36,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)산업공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(산업공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-26-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 14.86,
          "additionalPass": {
           "count": 18,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.64,
           "avgGradeReference90": 2.71,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)산업공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(산업공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-26-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 14.4,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.07,
           "avgGradeReference90": 3.25,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)산업공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(산업공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-26-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 17.5,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.87,
           "avgGradeReference90": 3.97,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)산업공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(산업공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-26-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 33.75,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.33,
           "avgGradeReference90": 4.61,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.3, 90%=11"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)산업공학전공\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(산업공학전공)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-27",
    "name": "생명과학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-27-schoolrecord-excellent",
     "comprehensive": "gachon-27-comprehensive-baram",
     "essay": "gachon-27-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-27-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 9.65,
          "additionalPass": {
           "count": 52,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.18,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-27-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 28.14,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.38,
           "avgGradeReference70": 2.78,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 6.4,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.41,
           "avgGradeReference90": 3.45,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 11.33,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.29,
           "avgGradeReference90": 2.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-27-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 2,
          "competitionRatio": 26,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.41,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 12.67,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.59,
           "avgGradeReference90": 2.72,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-27-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 77,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 2.85,
           "avgGradeReference70": 3.71,
           "avgGradeReference90": 3.81,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 51.5,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.93,
           "avgGradeReference90": 4.14,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 58.67,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.74,
           "avgGradeReference90": 3.79,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-27-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 39,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.82,
           "avgGradeReference90": 6.34,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107, 90%=106"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 43.2,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.36,
           "avgGradeReference90": 5.15,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=102.5, 90%=99"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 39.31,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.93,
           "avgGradeReference90": 4.53,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.4, 90%=11.1"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-28",
    "name": "설비·소방공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-28-schoolrecord-excellent",
     "comprehensive": "gachon-28-comprehensive-baram",
     "essay": "gachon-28-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-28-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 12.56,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.78,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)설비·소방공학과\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(설비·소방공학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-28-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 8,
          "competitionRatio": 16.38,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 3.13,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)설비·소방공학과\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(설비·소방공학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-28-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 21,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.51,
           "avgGradeReference90": 3.52,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)설비·소방공학과\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(설비·소방공학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-28-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 21.43,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.12,
           "avgGradeReference90": 4.42,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)설비·소방공학과\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(설비·소방공학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-28-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 25.64,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.68,
           "avgGradeReference90": 5.29,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.8, 90%=10.6"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(기계공학부)설비·소방공학과\" — 공식 자료상 기계공학부(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 기계공학부(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(설비·소방공학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-29",
    "name": "소프트웨어전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-29-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-29-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-29-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 35,
          "competitionRatio": 5.23,
          "additionalPass": {
           "count": 74,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 87.51,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-29-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 13,
          "competitionRatio": 9.15,
          "additionalPass": {
           "count": 42,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.67,
           "avgGradeReference90": 2.84,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-29-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 8,
          "competitionRatio": 8.38,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.19,
           "avgGradeReference90": 3.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": []
     },
     "essay": {
      "details": [
       {
        "id": "gachon-29-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 38.81,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.32,
           "avgGradeReference90": 4.99,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.8, 90%=11.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-30",
    "name": "스마트보안학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-30-regular-general1",
     "schoolRecord": "gachon-30-schoolrecord-excellent",
     "comprehensive": "gachon-30-comprehensive-baram",
     "essay": "gachon-30-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-30-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 17,
          "competitionRatio": 7.53,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.9,
           "percentileCut70": 86.6,
           "percentileCut90": 86.5,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 5.94,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.95,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "스마트보안전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-30-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 8,
          "competitionRatio": 8.63,
          "additionalPass": {
           "count": 22,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.87,
           "avgGradeReference70": 3.16,
           "avgGradeReference90": 3.28,
           "extra": [
            {
             "label": "2026학년도 모집단위명",
             "value": "정보보호학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 8,
          "competitionRatio": 7.6,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.94,
           "avgGradeReference90": 3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 2.99,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "스마트보안전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-30-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 10.5,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.34,
           "avgGradeReference90": 3.75,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "스마트보안전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-30-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 45.29,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.9,
           "avgGradeReference70": 4.05,
           "avgGradeReference90": 4.55,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "2026학년도 모집단위명",
             "value": "정보보호학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 18.9,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.75,
           "avgGradeReference90": 5.06,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 43.17,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.91,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "스마트보안전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-30-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 18,
          "competitionRatio": 28.17,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.88,
           "avgGradeReference90": 5.19,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=105, 90%=102"
            },
            {
             "label": "2026학년도 모집단위명",
             "value": "정보보호학과"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 17,
          "competitionRatio": 31.1,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.22,
           "avgGradeReference90": 6.08,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=101.5, 90%=100.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 33.38,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.49,
           "avgGradeReference90": 5,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.5, 90%=11.35"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "스마트보안전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-31",
    "name": "스마트시티학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-31-regular-general1",
     "schoolRecord": "gachon-31-schoolrecord-excellent",
     "comprehensive": "gachon-31-comprehensive-baram",
     "essay": "gachon-31-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-31-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 16,
          "competitionRatio": 5.75,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.2,
           "percentileCut70": 86.1,
           "percentileCut90": 86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 19,
          "competitionRatio": 7.53,
          "additionalPass": {
           "count": 46,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-31-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 8,
          "competitionRatio": 12.75,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.89,
           "avgGradeReference70": 2.98,
           "avgGradeReference90": 3.01,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 6.4,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.38,
           "avgGradeReference90": 3.44,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.71,
           "avgGradeReference90": 2.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-31-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 7,
          "competitionRatio": 15.7,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.04,
           "avgGradeReference90": 4.05,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 17.4,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.54,
           "avgGradeReference90": 3.86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-31-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 27.6,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.17,
           "avgGradeReference70": 4.23,
           "avgGradeReference90": 4.39,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 17.6,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.61,
           "avgGradeReference90": 4.78,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 28.67,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.93,
           "avgGradeReference90": 3.97,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-31-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 29.56,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.18,
           "avgGradeReference90": 4.86,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104, 90%=101"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 28.7,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.64,
           "avgGradeReference90": 6.31,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=98, 90%=97"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 30.58,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.22,
           "avgGradeReference90": 4.66,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.15, 90%=10.8"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-32",
    "name": "스마트팩토리전공",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-32-schoolrecord-excellent",
     "comprehensive": "gachon-32-comprehensive-baram",
     "essay": "gachon-32-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-32-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 18,
          "competitionRatio": 6.61,
          "additionalPass": {
           "count": 34,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-32-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 9,
          "competitionRatio": 6.9,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.22,
           "avgGradeReference90": 3.39,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.95,
           "avgGradeReference90": 2.99,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-32-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 14.8,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.66,
           "avgGradeReference90": 3.88,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-32-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.05,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 20.83,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.06,
           "avgGradeReference90": 4.15,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-32-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 31.1,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.95,
           "avgGradeReference90": 5.69,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=101, 90%=100"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 28.83,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.89,
           "avgGradeReference90": 5.73,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.9, 90%=10.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-33",
    "name": "스마트팩토리학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-33-regular-general1",
     "schoolRecord": "gachon-33-schoolrecord-excellent",
     "comprehensive": "gachon-33-comprehensive-baram",
     "essay": "gachon-33-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-33-regular-general1",
        "name": "일반전형1",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 18,
          "competitionRatio": 6.11,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.4,
           "percentileCut70": 86.2,
           "percentileCut90": 86.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-33-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 9,
          "competitionRatio": 13.78,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.92,
           "avgGradeReference70": 2.96,
           "avgGradeReference90": 2.97,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-33-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 15.29,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.04,
           "avgGradeReference70": 4.06,
           "avgGradeReference90": 4.22,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-33-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 28.63,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.8,
           "avgGradeReference90": 5.3,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=105"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-34",
    "name": "시스템반도체학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-34-regular-general1",
     "schoolRecord": "gachon-34-schoolrecord-excellent",
     "comprehensive": "gachon-34-comprehensive-baram",
     "essay": "gachon-34-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-34-regular-general1",
        "name": "일반전형1",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 18,
          "competitionRatio": 7.11,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 88.5,
           "percentileCut70": 88.3,
           "percentileCut90": 88,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-34-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 9,
          "competitionRatio": 17.78,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.99,
           "avgGradeReference70": 3.01,
           "avgGradeReference90": 3.12,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 9,
          "competitionRatio": 5.9,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.44,
           "avgGradeReference90": 3.47,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-34-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 13.57,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.44,
           "avgGradeReference70": 3.56,
           "avgGradeReference90": 3.6,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 9.6,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.78,
           "avgGradeReference90": 3.88,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-34-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 28.5,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.76,
           "avgGradeReference90": 6.11,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107, 90%=105.75"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 31.7,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.7,
           "avgGradeReference90": 7.04,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103, 90%=102"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-35",
    "name": "식품생명공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-35-regular-general1",
     "schoolRecord": "gachon-35-schoolrecord-excellent",
     "comprehensive": "gachon-35-comprehensive-baram",
     "essay": "gachon-35-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-35-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 16,
          "competitionRatio": 9.31,
          "additionalPass": {
           "count": 22,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.8,
           "percentileCut70": 86.7,
           "percentileCut90": 86.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 6.94,
          "additionalPass": {
           "count": 36,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-35-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 14.29,
          "additionalPass": {
           "count": 26,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.62,
           "avgGradeReference70": 2.7,
           "avgGradeReference90": 2.72,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 8.7,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.68,
           "avgGradeReference90": 2.95,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 10,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.56,
           "avgGradeReference90": 2.62,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-35-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.72,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 16.2,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 2.88,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-35-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 36.33,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.24,
           "avgGradeReference70": 3.56,
           "avgGradeReference90": 3.56,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 37.3,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.52,
           "avgGradeReference90": 3.96,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 38.67,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.8,
           "avgGradeReference90": 3.93,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-35-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 33.79,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.43,
           "avgGradeReference90": 5.26,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107, 90%=106"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 33.9,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.34,
           "avgGradeReference90": 5.45,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=106.25"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 34.29,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.19,
           "avgGradeReference90": 4.55,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.7, 90%=11.4"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-36",
    "name": "식품영양학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-36-regular-general1",
     "schoolRecord": "gachon-36-schoolrecord-excellent",
     "comprehensive": "gachon-36-comprehensive-baram",
     "essay": "gachon-36-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-36-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 14,
          "competitionRatio": 8.43,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 85.8,
           "percentileCut70": 85.8,
           "percentileCut90": 85.7,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 9.81,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.25,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-36-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 15.67,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.7,
           "avgGradeReference70": 2.75,
           "avgGradeReference90": 2.77,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 6,
          "competitionRatio": 10.8,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.02,
           "avgGradeReference90": 3.04,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 9.86,
          "additionalPass": {
           "count": 18,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.74,
           "avgGradeReference90": 2.82,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-36-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 19,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.84,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 10.83,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.3,
           "avgGradeReference90": 3.53,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-36-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 36.5,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.62,
           "avgGradeReference70": 3.86,
           "avgGradeReference90": 3.96,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 35.7,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.7,
           "avgGradeReference90": 3.94,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 32,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.89,
           "avgGradeReference90": 4.01,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-36-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 33.33,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.26,
           "avgGradeReference90": 5.19,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=106, 90%=104"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 30.1,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.88,
           "avgGradeReference90": 5.41,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=85.5, 90%=83.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 29.85,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.21,
           "avgGradeReference90": 4.73,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.2, 90%=11.1"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-37",
    "name": "신소재공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-37-regular-general1",
     "schoolRecord": "gachon-37-schoolrecord-excellent",
     "comprehensive": "gachon-37-comprehensive-baram",
     "essay": "gachon-37-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-37-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 16,
          "competitionRatio": 9.31,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.9,
           "percentileCut70": 87.5,
           "percentileCut90": 87.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 7.17,
          "additionalPass": {
           "count": 69,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-37-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 15.71,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.81,
           "avgGradeReference70": 2.92,
           "avgGradeReference90": 2.95,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 8.3,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.2,
           "avgGradeReference90": 3.25,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 11.57,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.65,
           "avgGradeReference90": 2.7,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-37-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 18,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.41,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.8,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.16,
           "avgGradeReference90": 3.32,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-37-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 20.67,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.18,
           "avgGradeReference70": 3.54,
           "avgGradeReference90": 3.59,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 14.5,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.52,
           "avgGradeReference90": 3.55,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 15.5,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.25,
           "avgGradeReference90": 3.49,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-37-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 35.57,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.69,
           "avgGradeReference90": 6.18,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=111, 90%=106"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 37.8,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.24,
           "avgGradeReference90": 6.28,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=93, 90%=90.25"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 36.46,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.03,
           "avgGradeReference90": 4.82,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.6, 90%=11.4"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-38",
    "name": "심리학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-38-regular-general1",
     "schoolRecord": "gachon-38-schoolrecord-excellent",
     "comprehensive": "gachon-38-comprehensive-baram",
     "essay": "gachon-38-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-38-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 15,
          "competitionRatio": 9.13,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.2,
           "percentileCut70": 86.3,
           "percentileCut90": 86.1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 18,
          "competitionRatio": 8.28,
          "additionalPass": {
           "count": 38,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.05,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-38-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 6,
          "competitionRatio": 11.67,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.35,
           "avgGradeReference70": 2.39,
           "avgGradeReference90": 2.47,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 11.8,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.35,
           "avgGradeReference90": 2.55,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.7,
           "avgGradeReference90": 2.71,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-38-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 1,
          "competitionRatio": 30,
          "additionalPass": {
           "count": "-",
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 13.83,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.01,
           "avgGradeReference90": 3.02,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-38-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 49,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.18,
           "avgGradeReference70": 3.28,
           "avgGradeReference90": 3.32,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 40.9,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.65,
           "avgGradeReference90": 3.77,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 69,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.46,
           "avgGradeReference90": 3.75,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-38-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 54.9,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.89,
           "avgGradeReference90": 6.56,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=112, 90%=111"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 47.4,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.82,
           "avgGradeReference90": 5.2,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110.25, 90%=109.25"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 39.8,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.94,
           "avgGradeReference90": 5,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.8, 90%=12.8"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-39",
    "name": "약학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-39-regular-general1",
     "schoolRecord": null,
     "comprehensive": null,
     "essay": null
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-39-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-fixed-2026",
          "recruitCount": 15,
          "competitionRatio": 6.47,
          "additionalPass": {
           "count": 26,
           "round": null
          },
          "admissionMethod": "수능 100%(국어25%+수학30%+영어20%+탐구2과목25% 고정비율, 의예과·약학과·한의예과 기준)",
          "indicators": {
           "percentileCut50": 96.8,
           "percentileCut70": 96.6,
           "percentileCut90": 96.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 15,
          "competitionRatio": 6.8,
          "additionalPass": {
           "count": 29,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 96.85,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-39-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 3,
          "competitionRatio": 13.33,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": "-",
           "avgGradeReference70": "-",
           "avgGradeReference90": 1.17,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 3,
          "competitionRatio": 37.7,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.13,
           "avgGradeReference90": 1.18,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 3,
          "competitionRatio": 36.33,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.2,
           "avgGradeReference90": 1.21,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-39-comprehensive-euiyak",
        "name": "가천의약학 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": []
     }
    }
   },
   {
    "id": "gachon-40",
    "name": "외국어계열",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-40-schoolrecord-excellent",
     "comprehensive": "gachon-40-comprehensive-baram",
     "essay": "gachon-40-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-40-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 63,
          "competitionRatio": 8.53,
          "additionalPass": {
           "count": 126,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.35,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)외국어계열\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(외국어계열)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-40-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 25,
          "competitionRatio": 15.2,
          "additionalPass": {
           "count": 64,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.89,
           "avgGradeReference90": 3.05,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)외국어계열\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(외국어계열)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-40-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 24,
          "competitionRatio": 17.46,
          "additionalPass": {
           "count": 35,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.27,
           "avgGradeReference90": 3.37,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)외국어계열\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(외국어계열)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-40-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 30,
          "competitionRatio": 25.5,
          "additionalPass": {
           "count": 23,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.21,
           "avgGradeReference90": 4.76,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)외국어계열\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(외국어계열)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-40-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 61,
          "competitionRatio": 24.3,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.84,
           "avgGradeReference90": 5.49,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.9, 90%=10.6"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)외국어계열\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(외국어계열)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-41",
    "name": "운동재활학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-41-regular-general1",
     "schoolRecord": "gachon-41-schoolrecord-excellent",
     "comprehensive": "gachon-41-comprehensive-baram",
     "essay": "gachon-41-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-41-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 14,
          "competitionRatio": 13.79,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.5,
           "percentileCut70": 86.4,
           "percentileCut90": 85.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 14,
          "competitionRatio": 8.93,
          "additionalPass": {
           "count": 38,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-41-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 41.33,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.87,
           "avgGradeReference70": 3.17,
           "avgGradeReference90": 3.23,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 5,
          "competitionRatio": 11.4,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4,
           "avgGradeReference90": 4.05,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 16.8,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.75,
           "avgGradeReference90": 2.86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-41-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 5,
          "competitionRatio": 27,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.2,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.45,
           "avgGradeReference90": 3.76,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-41-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 65.14,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.71,
           "avgGradeReference70": 3.92,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 53.3,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.33,
           "avgGradeReference90": 4.55,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 80.5,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4,
           "avgGradeReference90": 4.15,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-41-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 30.54,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.32,
           "avgGradeReference90": 5.92,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107, 90%=105"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 34.2,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.66,
           "avgGradeReference90": 5.01,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=98, 90%=97.75"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 29.7,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.48,
           "avgGradeReference90": 5.28,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.4, 90%=10"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-42",
    "name": "유아교육학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-42-regular-general1",
     "schoolRecord": "gachon-42-schoolrecord-excellent",
     "comprehensive": "gachon-42-comprehensive-baram",
     "essay": "gachon-42-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-42-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 13,
          "competitionRatio": 9.23,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 84.9,
           "percentileCut70": 84.7,
           "percentileCut90": 84.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 21,
          "competitionRatio": 6.45,
          "additionalPass": {
           "count": 54,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.25,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-42-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 6,
          "competitionRatio": 14.17,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 3.09,
           "avgGradeReference70": 3.15,
           "avgGradeReference90": 3.25,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 9.3,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.21,
           "avgGradeReference90": 3.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 10.5,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.78,
           "avgGradeReference90": 2.81,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-42-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-humanities",
          "recruitCount": 6,
          "competitionRatio": 27,
          "additionalPass": {
           "count": 4.16,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 3.13,
           "avgGradeReference70": 3.65,
           "avgGradeReference90": 3.66,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 28.5,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.66,
           "avgGradeReference90": 4.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 9.17,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.13,
           "avgGradeReference90": 3.46,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-42-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 33.63,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.51,
           "avgGradeReference70": 3.79,
           "avgGradeReference90": 3.86,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 31.3,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.93,
           "avgGradeReference90": 4.03,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 35.9,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.65,
           "avgGradeReference90": 3.72,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-42-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 40,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.36,
           "avgGradeReference90": 5.62,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=115, 90%=110"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 33.9,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.03,
           "avgGradeReference90": 5.71,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=109.5, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 15,
          "competitionRatio": 24.33,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.87,
           "avgGradeReference90": 5.6,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=10.8, 90%=10.2"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-43",
    "name": "응급구조학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-43-regular-general1",
     "schoolRecord": null,
     "comprehensive": "gachon-43-comprehensive-baram",
     "essay": "gachon-43-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-43-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 10,
          "competitionRatio": 11.4,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 85.5,
           "percentileCut70": 85.1,
           "percentileCut90": 85.1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 10,
          "competitionRatio": 6.3,
          "additionalPass": {
           "count": 23,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.54,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": []
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-43-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 4,
          "competitionRatio": 83,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 4.13,
           "avgGradeReference70": 4.64,
           "avgGradeReference90": 4.72,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 4,
          "competitionRatio": 116.5,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.11,
           "avgGradeReference90": 4.35,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 4,
          "competitionRatio": 66.75,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 5.1,
           "avgGradeReference90": 6.07,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-43-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 39.83,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.84,
           "avgGradeReference90": 5.41,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103, 90%=102"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 28.8,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.86,
           "avgGradeReference90": 5.78,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=81.5, 90%=80"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 28,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.17,
           "avgGradeReference90": 4.23,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.4, 90%=10.6"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-44",
    "name": "응용통계학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-44-regular-general1",
     "schoolRecord": "gachon-44-schoolrecord-excellent",
     "comprehensive": "gachon-44-comprehensive-baram",
     "essay": "gachon-44-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-44-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 16,
          "competitionRatio": 7.25,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.3,
           "percentileCut70": 86.2,
           "percentileCut90": 85.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 16,
          "competitionRatio": 6.69,
          "additionalPass": {
           "count": 29,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-44-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 7,
          "competitionRatio": 9.71,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.77,
           "avgGradeReference70": 2.85,
           "avgGradeReference90": 2.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 7,
          "competitionRatio": 9.4,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.73,
           "avgGradeReference90": 2.87,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 11.17,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.79,
           "avgGradeReference90": 2.91,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-44-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 2,
          "competitionRatio": 21,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 7,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.85,
           "avgGradeReference90": 4.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-44-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 14.71,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.51,
           "avgGradeReference70": 3.55,
           "avgGradeReference90": 3.58,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.12,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 13.29,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.71,
           "avgGradeReference90": 4.12,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-44-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 41.38,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.28,
           "avgGradeReference90": 5.93,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=116, 90%=114.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 34.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.03,
           "avgGradeReference90": 5.92,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=114, 90%=113"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 31.54,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.5,
           "avgGradeReference90": 4.7,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.8, 90%=11.3"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-45",
    "name": "의공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-45-regular-general1",
     "schoolRecord": "gachon-45-schoolrecord-excellent",
     "comprehensive": "gachon-45-comprehensive-baram",
     "essay": "gachon-45-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-45-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 17,
          "competitionRatio": 8.59,
          "additionalPass": {
           "count": 26,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.6,
           "percentileCut70": 86.4,
           "percentileCut90": 86.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 19,
          "competitionRatio": 9.74,
          "additionalPass": {
           "count": 45,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.65,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-45-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 6,
          "competitionRatio": 32.33,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.82,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 2.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 6,
          "competitionRatio": 8.7,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.83,
           "avgGradeReference90": 4.13,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 12.83,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.46,
           "avgGradeReference90": 2.55,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-45-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 12,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 4.29,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 8.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.03,
           "avgGradeReference90": 3.1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-45-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 19.33,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.51,
           "avgGradeReference70": 3.61,
           "avgGradeReference90": 3.73,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 11.8,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.92,
           "avgGradeReference90": 4.06,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 15.5,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.33,
           "avgGradeReference90": 3.43,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-45-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 26.14,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.11,
           "avgGradeReference90": 5.81,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=102, 90%=98"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 30.2,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.5,
           "avgGradeReference90": 6.06,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103.5, 90%=98.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 27.17,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.79,
           "avgGradeReference90": 5.46,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.2, 90%=11"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-46",
    "name": "의료산업경영학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-46-regular-general1",
     "schoolRecord": "gachon-46-schoolrecord-excellent",
     "comprehensive": "gachon-46-comprehensive-baram",
     "essay": "gachon-46-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-46-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 17,
          "competitionRatio": 10.12,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.5,
           "percentileCut70": 86.5,
           "percentileCut90": 86.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 9.35,
          "additionalPass": {
           "count": 46,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.45,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-46-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 7,
          "competitionRatio": 13.43,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.56,
           "avgGradeReference70": 2.74,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 7,
          "competitionRatio": 8.1,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.14,
           "avgGradeReference90": 3.18,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15.83,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.63,
           "avgGradeReference90": 2.84,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-46-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 1,
          "competitionRatio": 23,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 5.27,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.07,
           "avgGradeReference90": 3.22,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-46-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.75,
           "avgGradeReference90": 4.11,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 23,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.76,
           "avgGradeReference90": 3.92,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-46-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 41.08,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5,
           "avgGradeReference90": 5.77,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=113, 90%=110.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 33.8,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.85,
           "avgGradeReference90": 5.96,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=111, 90%=108"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 29.08,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.65,
           "avgGradeReference90": 5.02,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.6, 90%=11.3"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-47",
    "name": "의예과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-47-regular-general1",
     "schoolRecord": "gachon-47-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-47-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-47-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-fixed-2026",
          "recruitCount": 13,
          "competitionRatio": 7.92,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "수능 100%(국어25%+수학30%+영어20%+탐구2과목25% 고정비율, 의예과·약학과·한의예과 기준)",
          "indicators": {
           "percentileCut50": 98.6,
           "percentileCut70": 98.5,
           "percentileCut90": 98.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 15,
          "competitionRatio": 5.93,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 98.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-47-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 4,
          "competitionRatio": 15.75,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 1,
           "avgGradeReference70": 1.04,
           "avgGradeReference90": 1.06,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 15,
          "competitionRatio": 13.4,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.05,
           "avgGradeReference90": 1.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 25.2,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1,
           "avgGradeReference90": 1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-47-comprehensive-euiyak",
        "name": "가천의약학 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-47-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 577,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.26,
           "avgGradeReference90": 3.63,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=123, 90%=118"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 40,
          "competitionRatio": 205.2,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.31,
           "avgGradeReference90": 4,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=78.75, 90%=77.25"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-48",
    "name": "인공지능학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-48-regular-general1",
     "schoolRecord": "gachon-48-schoolrecord-excellent",
     "comprehensive": "gachon-48-comprehensive-baram",
     "essay": "gachon-48-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-48-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 46,
          "competitionRatio": 7.52,
          "additionalPass": {
           "count": 45,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.8,
           "percentileCut70": 87.4,
           "percentileCut90": 87.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 51,
          "competitionRatio": 6.84,
          "additionalPass": {
           "count": 103,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 85.44,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "인공지능전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-48-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 21,
          "competitionRatio": 17.71,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 130.7,
           "standardScoreCut70": 129.9,
           "standardScoreCut90": 129.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-48-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 21,
          "competitionRatio": 7.86,
          "additionalPass": {
           "count": 40,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.72,
           "avgGradeReference70": 2.8,
           "avgGradeReference90": 2.94,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 20,
          "competitionRatio": 7.2,
          "additionalPass": {
           "count": 32,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.81,
           "avgGradeReference90": 2.86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 16,
          "competitionRatio": 10.5,
          "additionalPass": {
           "count": 44,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.83,
           "avgGradeReference90": 2.86,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "인공지능전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-48-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 15,
          "competitionRatio": 14.4,
          "additionalPass": {
           "count": 28,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.22,
           "avgGradeReference90": 3.25,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "인공지능전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-48-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 25.38,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.62,
           "avgGradeReference70": 3.74,
           "avgGradeReference90": 3.95,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 18.1,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.59,
           "avgGradeReference90": 4.07,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-48-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 45,
          "competitionRatio": 33.6,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.16,
           "avgGradeReference90": 7.18,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=107, 90%=106"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 40,
          "competitionRatio": 35.2,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.32,
           "avgGradeReference90": 5.8,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104, 90%=103.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 35,
          "competitionRatio": 35.37,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.48,
           "avgGradeReference90": 5.21,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.8, 90%=11.5"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "인공지능전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-49",
    "name": "자유전공",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-49-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-49-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-49-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 21,
          "competitionRatio": 8.57,
          "additionalPass": {
           "count": 65,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.67,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-49-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 7,
          "competitionRatio": 16.43,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.5,
           "avgGradeReference90": 2.51,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-49-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-humanities",
          "recruitCount": 321,
          "competitionRatio": 15.88,
          "additionalPass": {
           "count": 4.07,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 3.01,
           "avgGradeReference70": 3.18,
           "avgGradeReference90": 3.45,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 321,
          "competitionRatio": 20.5,
          "additionalPass": {
           "count": 107,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.48,
           "avgGradeReference90": 5.14,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 5,
          "competitionRatio": 28.6,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.2,
           "avgGradeReference90": 3.27,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": []
     },
     "essay": {
      "details": [
       {
        "id": "gachon-49-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 18,
          "competitionRatio": 33.56,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.39,
           "avgGradeReference90": 5.25,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.2, 90%=11.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-50",
    "name": "전기공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-50-regular-general1",
     "schoolRecord": "gachon-50-schoolrecord-excellent",
     "comprehensive": "gachon-50-comprehensive-baram",
     "essay": "gachon-50-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-50-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 21,
          "competitionRatio": 6.71,
          "additionalPass": {
           "count": 20,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.4,
           "percentileCut70": 87.2,
           "percentileCut90": 86.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 29,
          "competitionRatio": 5.72,
          "additionalPass": {
           "count": 65,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 85.32,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-50-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 8,
          "competitionRatio": 15.38,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.96,
           "avgGradeReference70": 3,
           "avgGradeReference90": 3.11,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 8,
          "competitionRatio": 7.3,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.25,
           "avgGradeReference90": 3.35,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 9,
          "competitionRatio": 15,
          "additionalPass": {
           "count": 27,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.72,
           "avgGradeReference90": 2.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-50-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.18,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 7,
          "competitionRatio": 17.14,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.25,
           "avgGradeReference90": 3.33,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-50-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 16.89,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.37,
           "avgGradeReference70": 3.45,
           "avgGradeReference90": 3.93,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 21.2,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.88,
           "avgGradeReference90": 3.97,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.24,
           "avgGradeReference90": 4.31,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-50-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 20,
          "competitionRatio": 30.75,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.18,
           "avgGradeReference90": 5.83,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=110.5, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 20,
          "competitionRatio": 34.1,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.09,
           "avgGradeReference90": 5.25,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104, 90%=103"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 22,
          "competitionRatio": 33.18,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.58,
           "avgGradeReference90": 5.51,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.2, 90%=10.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-51",
    "name": "치위생학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-51-regular-general1",
     "schoolRecord": "gachon-51-schoolrecord-excellent",
     "comprehensive": "gachon-51-comprehensive-baram",
     "essay": "gachon-51-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-51-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 15,
          "competitionRatio": 11.8,
          "additionalPass": {
           "count": 31,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 85.6,
           "percentileCut70": 85,
           "percentileCut90": 84.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 15,
          "competitionRatio": 6,
          "additionalPass": {
           "count": 37,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 79.84,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-51-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 5,
          "competitionRatio": 15.6,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.81,
           "avgGradeReference70": 2.86,
           "avgGradeReference90": 3.08,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 5,
          "competitionRatio": 6.2,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.84,
           "avgGradeReference90": 3.96,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.2,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.54,
           "avgGradeReference90": 2.63,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-51-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-balanced-2026-science",
          "recruitCount": 5,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 4.16,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": 3.17,
           "avgGradeReference70": 3.48,
           "avgGradeReference90": 3.69,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 5,
          "competitionRatio": 17.8,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.99,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 7.8,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.25,
           "avgGradeReference90": 3.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-51-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 22.5,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.3,
           "avgGradeReference70": 3.67,
           "avgGradeReference90": 3.89,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.9,
           "avgGradeReference90": 4.09,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 19.33,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.5,
           "avgGradeReference90": 3.5,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-51-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 32,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.3,
           "avgGradeReference90": 5.92,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=104, 90%=101"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 31.3,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.13,
           "avgGradeReference90": 4.44,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=78.5, 90%=78"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 9,
          "competitionRatio": 28.56,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.51,
           "avgGradeReference90": 5.12,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.5, 90%=11.3"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-52",
    "name": "컴퓨터공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-52-regular-general1",
     "schoolRecord": "gachon-52-schoolrecord-excellent",
     "comprehensive": "gachon-52-comprehensive-baram",
     "essay": "gachon-52-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-52-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 40,
          "competitionRatio": 6.2,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 87.4,
           "percentileCut70": 87,
           "percentileCut90": 86.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 55,
          "competitionRatio": 5.71,
          "additionalPass": {
           "count": 105,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 86.46,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "컴퓨터공학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-52-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 18,
          "competitionRatio": 13.17,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.2,
           "standardScoreCut70": 128.7,
           "standardScoreCut90": 128.6,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-52-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 16,
          "competitionRatio": 7.81,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.6,
           "avgGradeReference70": 2.64,
           "avgGradeReference90": 2.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 20,
          "competitionRatio": 7.1,
          "additionalPass": {
           "count": 28,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.74,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 18,
          "competitionRatio": 11.94,
          "additionalPass": {
           "count": 41,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.53,
           "avgGradeReference90": 2.58,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "컴퓨터공학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-52-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 15,
          "competitionRatio": 8.67,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.29,
           "avgGradeReference90": 3.39,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "컴퓨터공학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-52-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 32.57,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.33,
           "avgGradeReference70": 3.46,
           "avgGradeReference90": 3.65,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 21,
          "competitionRatio": 25.7,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.81,
           "avgGradeReference90": 4.1,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 22,
          "competitionRatio": 23.45,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.57,
           "avgGradeReference90": 3.9,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "컴퓨터공학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-52-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 41,
          "competitionRatio": 33.68,
          "additionalPass": {
           "count": 12,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.93,
           "avgGradeReference90": 5.55,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=109, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 40,
          "competitionRatio": 46.5,
          "additionalPass": {
           "count": 14,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.25,
           "avgGradeReference90": 6.18,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=106.5, 90%=105"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 35,
          "competitionRatio": 48.31,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.15,
           "avgGradeReference90": 5,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.7, 90%=12.5"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "컴퓨터공학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-53",
    "name": "클라우드공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-53-regular-general1",
     "schoolRecord": "gachon-53-schoolrecord-excellent",
     "comprehensive": "gachon-53-comprehensive-baram",
     "essay": "gachon-53-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-53-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 11,
          "competitionRatio": 5.55,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 90.4,
           "percentileCut70": 89.9,
           "percentileCut90": 89.7,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 9,
          "competitionRatio": 11.22,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 90.95,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-53-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 21.43,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.44,
           "avgGradeReference70": 2.54,
           "avgGradeReference90": 2.62,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 6.6,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.25,
           "avgGradeReference90": 3.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 27.86,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.06,
           "avgGradeReference90": 2.33,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-53-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.38,
           "avgGradeReference70": 3.42,
           "avgGradeReference90": 3.55,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 10.9,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.58,
           "avgGradeReference90": 3.59,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 21.71,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.64,
           "avgGradeReference90": 2.95,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-53-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 77.29,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.44,
           "avgGradeReference90": 5.67,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=127, 90%=124"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 86,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.32,
           "avgGradeReference90": 5.83,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=115.5, 90%=111"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 135,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.04,
           "avgGradeReference90": 3.41,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=12.3, 90%=12.3"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-54",
    "name": "토목환경공학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-54-schoolrecord-excellent",
     "comprehensive": "gachon-54-comprehensive-baram",
     "essay": "gachon-54-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-54-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 7.33,
          "additionalPass": {
           "count": 52,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 83.53,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-54-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.84,
           "avgGradeReference90": 2.99,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 7,
          "competitionRatio": 19.86,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.02,
           "avgGradeReference90": 3.04,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-54-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 1,
          "competitionRatio": 14,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.74,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 5,
          "competitionRatio": 12.8,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.4,
           "avgGradeReference90": 3.44,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-54-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 26.8,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.68,
           "avgGradeReference90": 3.88,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 23.17,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.95,
           "avgGradeReference90": 3.97,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-54-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 30.5,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.93,
           "avgGradeReference90": 5.93,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=102, 90%=98.75"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 29.08,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.73,
           "avgGradeReference90": 5.27,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=9.8, 90%=9.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-55",
    "name": "패션산업학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-55-regular-general1",
     "schoolRecord": "gachon-55-schoolrecord-excellent",
     "comprehensive": "gachon-55-comprehensive-baram",
     "essay": "gachon-55-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-55-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 13,
          "competitionRatio": 11,
          "additionalPass": {
           "count": 24,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 85.9,
           "percentileCut70": 85.7,
           "percentileCut90": 85.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 13,
          "competitionRatio": 10.31,
          "additionalPass": {
           "count": 24,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 82.15,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-55-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 6,
          "competitionRatio": 16.67,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.66,
           "avgGradeReference70": 2.75,
           "avgGradeReference90": 2.88,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 12.7,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.09,
           "avgGradeReference90": 3.26,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 15,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3,
           "avgGradeReference90": 3.1,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-55-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-humanities",
          "recruitCount": 6,
          "competitionRatio": 24,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.12,
           "avgGradeReference90": 4.25,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 5,
          "competitionRatio": 15.8,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.26,
           "avgGradeReference90": 3.43,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-55-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 42.43,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.79,
           "avgGradeReference70": 3.9,
           "avgGradeReference90": 4.37,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 50.2,
          "additionalPass": {
           "count": 6,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.06,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 6,
          "competitionRatio": 55.83,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.24,
           "avgGradeReference90": 4.26,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-55-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 44.42,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.92,
           "avgGradeReference90": 6.28,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=120, 90%=117"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 50.7,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.33,
           "avgGradeReference90": 4.43,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=118.5, 90%=117.25"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 38.38,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 3.98,
           "avgGradeReference90": 4.96,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.9, 90%=11.7"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-56",
    "name": "한국어문학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-56-schoolrecord-excellent",
     "comprehensive": "gachon-56-comprehensive-baram",
     "essay": "gachon-56-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-56-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 6.56,
          "additionalPass": {
           "count": 28,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.6,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)한국어문학과\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(한국어문학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-56-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 13.67,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.88,
           "avgGradeReference90": 2.97,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)한국어문학과\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(한국어문학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-56-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 9.83,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.38,
           "avgGradeReference90": 3.4,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)한국어문학과\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(한국어문학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-56-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 8,
          "competitionRatio": 29.13,
          "additionalPass": {
           "count": 5,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.96,
           "avgGradeReference90": 4.32,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)한국어문학과\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(한국어문학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-56-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 11,
          "competitionRatio": 26.55,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.45,
           "avgGradeReference90": 5.53,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.7, 90%=11"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(AI인문대학)한국어문학과\" — 공식 자료상 AI인문대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 AI인문대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(한국어문학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-57",
    "name": "한의예과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-57-regular-general1",
     "schoolRecord": "gachon-57-schoolrecord-excellent",
     "comprehensive": null,
     "essay": null
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-57-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-fixed-2026",
          "recruitCount": 21,
          "competitionRatio": 8.24,
          "additionalPass": {
           "count": 36,
           "round": null
          },
          "admissionMethod": "수능 100%(국어25%+수학30%+영어20%+탐구2과목25% 고정비율, 의예과·약학과·한의예과 기준)",
          "indicators": {
           "percentileCut50": 97.6,
           "percentileCut70": 97.5,
           "percentileCut90": 97.4,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 15,
          "competitionRatio": 7.57,
          "additionalPass": {
           "count": 45,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 98.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-57-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 19.43,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 1.11,
           "avgGradeReference70": 1.15,
           "avgGradeReference90": 1.21,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 5,
          "competitionRatio": 27.4,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.12,
           "avgGradeReference90": 1.27,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 5,
          "competitionRatio": 39,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 1.12,
           "avgGradeReference90": 1.13,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-57-comprehensive-euiyak",
        "name": "가천의약학 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": null,
          "competitionRatio": null,
          "additionalPass": {
           "count": null,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "extra": [
            {
             "label": "결과 공시",
             "value": "입시결과 파일에 가천의약학 전형 결과 행이 없어 확인 불가(의약학 계열 결과 비공개 정책으로 추정)"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026가천대학교수시모집요강.pdf",
            "page": 58
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": []
     }
    }
   },
   {
    "id": "gachon-58",
    "name": "행정학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-58-schoolrecord-excellent",
     "comprehensive": "gachon-58-comprehensive-baram",
     "essay": "gachon-58-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-58-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "가군",
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 18,
          "competitionRatio": 10.56,
          "additionalPass": {
           "count": 47,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 81.85,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-58-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 6,
          "competitionRatio": 16,
          "additionalPass": {
           "count": 29,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.76,
           "avgGradeReference90": 2.77,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-58-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 7,
          "competitionRatio": 9.57,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.27,
           "avgGradeReference90": 3.44,
           "extra": [
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-58-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 20.71,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.5,
           "avgGradeReference90": 3.76,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-58-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 10,
          "competitionRatio": 29.8,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.57,
           "avgGradeReference90": 5.11,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11, 90%=10.9"
            },
            {
             "label": "모집단위 개편 참고",
             "value": "원본 표기 \"(법과대학)행정학과\" — 공식 자료상 법과대학(으)로 통합될 예정이나, 같은 시기에 다른 모집단위도 함께 법과대학(으)로 통합되어(1:N 관계) 어느 비중으로 이어지는지 특정할 수 없으므로 임의로 병합하지 않고 이 학년도의 실제 모집단위(행정학과)를 그대로 별개 모집단위로 유지함"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-59",
    "name": "화공생명배터리공학부",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-59-regular-general1",
     "schoolRecord": "gachon-59-schoolrecord-excellent",
     "comprehensive": "gachon-59-comprehensive-baram",
     "essay": "gachon-59-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-59-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 61,
          "competitionRatio": 5.98,
          "additionalPass": {
           "count": 40,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.6,
           "percentileCut70": 86.5,
           "percentileCut90": 86.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 81,
          "competitionRatio": 5.46,
          "additionalPass": {
           "count": 98,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.9,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-59-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 25,
          "competitionRatio": 14.32,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.2,
           "standardScoreCut70": 129,
           "standardScoreCut90": 128.8,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-59-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 25,
          "competitionRatio": 7.56,
          "additionalPass": {
           "count": 32,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.64,
           "avgGradeReference70": 2.73,
           "avgGradeReference90": 2.78,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 30,
          "competitionRatio": 6,
          "additionalPass": {
           "count": 33,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.77,
           "avgGradeReference90": 2.83,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 25,
          "competitionRatio": 10.64,
          "additionalPass": {
           "count": 63,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.66,
           "avgGradeReference90": 2.86,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-59-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 15,
          "competitionRatio": 11.93,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.08,
           "avgGradeReference90": 3.26,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-59-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 30,
          "competitionRatio": 18.7,
          "additionalPass": {
           "count": 10,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.66,
           "avgGradeReference70": 3.86,
           "avgGradeReference90": 4.13,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 36,
          "competitionRatio": 12.9,
          "additionalPass": {
           "count": 21,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.7,
           "avgGradeReference90": 4.15,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 28,
          "competitionRatio": 22.25,
          "additionalPass": {
           "count": 13,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.5,
           "avgGradeReference90": 3.8,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-59-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 57,
          "competitionRatio": 29.65,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.5,
           "avgGradeReference90": 5.26,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=113, 90%=111.5"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 58,
          "competitionRatio": 29.7,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.47,
           "avgGradeReference90": 5.26,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=92, 90%=90"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 46,
          "competitionRatio": 30.35,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.26,
           "avgGradeReference90": 4.94,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.7, 90%=11.4"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-60",
    "name": "화학과",
    "college": null,
    "track": "science",
    "simulationBinding": {
     "regular": "gachon-60-regular-general1",
     "schoolRecord": "gachon-60-schoolrecord-excellent",
     "comprehensive": "gachon-60-comprehensive-baram",
     "essay": "gachon-60-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-60-regular-general1",
        "name": "일반전형(2026학년도부터 일반전형1)",
        "group": "다군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 15,
          "competitionRatio": 6.07,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.4,
           "percentileCut70": 86.2,
           "percentileCut90": 86.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-regular-general1-2024",
          "recruitCount": 17,
          "competitionRatio": 6.76,
          "additionalPass": {
           "count": 35,
           "round": null
          },
          "admissionMethod": "수능 100%(2024학년도 반영비율은 hwp 원문 파싱 불가로 확인 못함)",
          "indicators": {
           "percentileCut70": 84.18,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2024학년도정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-60-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-science",
          "recruitCount": 7,
          "competitionRatio": 11.14,
          "additionalPass": {
           "count": 7,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.65,
           "avgGradeReference70": 2.69,
           "avgGradeReference90": 2.71,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-science",
          "recruitCount": 7,
          "competitionRatio": 15.7,
          "additionalPass": {
           "count": 8,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.74,
           "avgGradeReference90": 2.79,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-science",
          "recruitCount": 6,
          "competitionRatio": 9.33,
          "additionalPass": {
           "count": 15,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.22,
           "avgGradeReference90": 3.5,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-60-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-balanced-2025-science",
          "recruitCount": 2,
          "competitionRatio": 18.5,
          "additionalPass": {
           "count": 0,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": null,
           "avgGradeReference90": 3.74,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-science",
          "recruitCount": 6,
          "competitionRatio": 8.83,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.33,
           "avgGradeReference90": 3.44,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-60-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 24.86,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.65,
           "avgGradeReference70": 3.67,
           "avgGradeReference90": 4,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 31.3,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.52,
           "avgGradeReference90": 3.58,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 7,
          "competitionRatio": 29,
          "additionalPass": {
           "count": 3,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.86,
           "avgGradeReference90": 4.21,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-60-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 14,
          "competitionRatio": 27.64,
          "additionalPass": {
           "count": 2,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.79,
           "avgGradeReference90": 5.79,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 28.3,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.71,
           "avgGradeReference90": 6.05,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=103, 90%=102"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 12,
          "competitionRatio": 33.58,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.07,
           "avgGradeReference90": 4.8,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.2, 90%=11"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-61",
    "name": "회계세무학과",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": "gachon-61-regular-general1",
     "schoolRecord": "gachon-61-schoolrecord-excellent",
     "comprehensive": null,
     "essay": "gachon-61-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-61-regular-general1",
        "name": "일반전형1",
        "group": "나군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general1-2026",
          "recruitCount": 15,
          "competitionRatio": 7.67,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어 중 우수 영역순 35%·25%·20% + 탐구 1과목 20%)",
          "indicators": {
           "percentileCut50": 86.8,
           "percentileCut70": 86.7,
           "percentileCut90": 86.5,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-61-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 7,
          "competitionRatio": 12.43,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.67,
           "avgGradeReference70": 2.67,
           "avgGradeReference90": 2.73,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 7,
          "competitionRatio": 11.7,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.9,
           "avgGradeReference90": 2.92,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-excellent-2024-humanities",
          "recruitCount": 10,
          "competitionRatio": 22.1,
          "additionalPass": {
           "count": 19,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 2.93,
           "avgGradeReference90": 2.93,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "회계세무학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       },
       {
        "id": "gachon-61-schoolrecord-balanced",
        "name": "지역균형 전형",
        "group": null,
        "years": [
         {
          "year": 2024,
          "calculationRuleId": "gachon-schoolrecord-balanced-2024-humanities",
          "recruitCount": 7,
          "competitionRatio": 7.86,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "1단계 학생부교과 100%(7배수) → 2단계 1단계 50% + 면접 50%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.45,
           "avgGradeReference90": 3.54,
           "extra": [
            {
             "label": "2024학년도 모집단위명",
             "value": "회계세무학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": []
     },
     "essay": {
      "details": [
       {
        "id": "gachon-61-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 42.06,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.45,
           "avgGradeReference90": 4.8,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108, 90%=107"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 16,
          "competitionRatio": 36.9,
          "additionalPass": {
           "count": 1,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.35,
           "avgGradeReference90": 6.27,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=111, 90%=110"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2024,
          "calculationRuleId": null,
          "recruitCount": 13,
          "competitionRatio": 31.92,
          "additionalPass": {
           "count": 4,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 4.61,
           "avgGradeReference90": 5,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=11.3, 90%=11.1"
            },
            {
             "label": "2024학년도 모집단위명",
             "value": "회계세무학전공"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "가천대학교_2024학년도_수시_입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   },
   {
    "id": "gachon-62",
    "name": "AI인문대학",
    "college": null,
    "track": "humanities",
    "simulationBinding": {
     "regular": null,
     "schoolRecord": "gachon-62-schoolrecord-excellent",
     "comprehensive": "gachon-62-comprehensive-baram",
     "essay": "gachon-62-essay-general"
    },
    "admissionTypes": {
     "regular": {
      "details": [
       {
        "id": "gachon-62-regular-general2",
        "name": "일반전형2(2026학년도 신설)",
        "group": "가군",
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-regular-general2-2026",
          "recruitCount": 26,
          "competitionRatio": 22.88,
          "additionalPass": {
           "count": 17,
           "round": null
          },
          "admissionMethod": "수능 100%(국어·수학·영어·탐구 1과목 중 우수 영역순 50%·30%·20%, 최하위 1개 영역 미반영)",
          "indicators": {
           "standardScoreCut50": 129.7,
           "standardScoreCut70": 129.4,
           "standardScoreCut90": 129.3,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026정시입시결과.xls"
           }
          ]
         }
        ]
       }
      ]
     },
     "schoolRecord": {
      "details": [
       {
        "id": "gachon-62-schoolrecord-excellent",
        "name": "학생부우수자 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": "gachon-schoolrecord-excellent-2026-humanities",
          "recruitCount": 22,
          "competitionRatio": 14.59,
          "additionalPass": {
           "count": 62,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": 2.87,
           "avgGradeReference70": 2.9,
           "avgGradeReference90": 3.01,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": "gachon-schoolrecord-excellent-2025-humanities",
          "recruitCount": 30,
          "competitionRatio": 8.8,
          "additionalPass": {
           "count": 58,
           "round": null
          },
          "admissionMethod": "학생부교과 100%",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 3.06,
           "avgGradeReference90": 3.2,
           "extra": []
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "comprehensive": {
      "details": [
       {
        "id": "gachon-62-comprehensive-baram",
        "name": "가천바람개비 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 25,
          "competitionRatio": 34.36,
          "additionalPass": {
           "count": 16,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": 3.78,
           "avgGradeReference70": 3.95,
           "avgGradeReference90": 4.96,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 30,
          "competitionRatio": 25.3,
          "additionalPass": {
           "count": 23,
           "round": null
          },
          "admissionMethod": "1단계 서류 100%(5배수) → 2단계 1단계 50% + 면접 50%(정성평가, 정량화 없음)",
          "indicators": {
           "avgGradeReference50": null,
           "avgGradeReference70": 4.28,
           "avgGradeReference90": 4.9,
           "extra": [
            {
             "label": "평가요소(서류·면접 공통)",
             "value": "인성 40% · 진학의지 및 계열적합성 40% · 학업역량(면접은 의사소통역량) 20%"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     },
     "essay": {
      "details": [
       {
        "id": "gachon-62-essay-general",
        "name": "논술 전형",
        "group": null,
        "years": [
         {
          "year": 2026,
          "calculationRuleId": null,
          "recruitCount": 71,
          "competitionRatio": 45.2,
          "additionalPass": {
           "count": 9,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.29,
           "avgGradeReference90": 6.16,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=117, 90%=115"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2026학년도수시입시결과.xlsx"
           }
          ]
         },
         {
          "year": 2025,
          "calculationRuleId": null,
          "recruitCount": 71,
          "competitionRatio": 34.6,
          "additionalPass": {
           "count": 11,
           "round": null
          },
          "admissionMethod": "논술 100%(인문 국어9/수학6문항, 자연 국어6/수학9문항, 각 150점+기본점수850점)",
          "indicators": {
           "avgGradeReference70": 5.03,
           "avgGradeReference90": 5.84,
           "extra": [
            {
             "label": "논술 원점수 컷(150점 만점)",
             "value": "70%=108.75, 90%=106"
            }
           ]
          },
          "sources": [
           {
            "org": "가천대학교 입학처(사용자 제공 자료)",
            "url": null,
            "publishedDate": null,
            "checkedDate": "2026-09-01",
            "title": "2025학년도수시입시결과.xlsx"
           }
          ]
         }
        ]
       }
      ]
     }
    }
   }
  ];
  UNIVERSITY_DB.push({id:'gachon',name:'가천대학교',aliases:['가천대'],region:'경기',departments});
})();
