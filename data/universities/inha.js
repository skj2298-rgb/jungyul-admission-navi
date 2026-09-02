// data/universities/inha.js
// 인하대학교 — 대학 구조 DB(UNIVERSITY_DB) 등록 파일.
// classic <script src> 로 로드되며 UNIVERSITY_DB/CALCULATION_RULES 전역에 push/할당만 한다 (fetch·import 없음, file:// 직접 열기 호환).
//
// 정시: 기존 index.html에 있던 INHA_2025_RESULT_ROWS(대학어디가 공개표준안 2025학년도 인하대 수능일반 결과, 60개 모집단위)를
// 수치 변경 없이 그대로 이 파일로 옮기고, 새 UNIVERSITY_DB 구조(대학→모집단위→전형유형→세부전형→학년도)로 재구성한다.
// 실제 학생 성적 대비 계산은 인하대 전용 백분위 환산 엔진(index.html의 inhaReferenceScore/inhaProfileReferenceScore)을 그대로 사용하며,
// 이 파일의 CALCULATION_RULES 항목은 그 반영 구조를 문서화하는 참조용 레코드다.
(function(){
  const INHA_SOURCE_URL='https://www.adiga.kr/ucp/uvt/uni/univDetailSelection.do?menuId=PCUVTINF2000&searchSyr=2026&unvCd=0000169';
  const INHA_2027_GUIDE_URL='https://www.adiga.kr/ucp/uvt/uni/univDetailSelection.do?menuId=PCUVTINF2000&searchSyr=2027&unvCd=0000169';

  // 계열별 반영비율 옵션(기존 INHA_WEIGHT_OPTIONS와 동일)
  const WEIGHT_OPTIONS_BY_TRACK={
    H:[{kor:35,mat:25,eng:15,inq:25}],
    N:[{kor:25,mat:35,eng:10,inq:30}],
    FH:[{kor:35,mat:30,eng:10,inq:25},{kor:30,mat:30,eng:10,inq:30}],
    FN:[{kor:25,mat:40,eng:10,inq:25},{kor:25,mat:35,eng:10,inq:30}]
  };

  // 대학어디가 공개표준안의 2025학년도 인하대학교 수능(일반) 결과 원자료 (기존 index.html INHA_2025_RESULT_ROWS와 동일, 수치 변경 없음)
  // 배열: 모집단위, 군, 최종모집, 경쟁률, 충원합격순위, 환산50컷, 환산70컷, 계열, 70% 등록자 [국,수,탐1,탐2,영,한국사], 2025결과표 모집단위명(개편 시)
  const INHA_2025_RESULT_ROWS=[
    ['조선해양공학과','가군',7,6.0,3,890.8,890.6,'N',[72,81,90,97,2,1]],
    ['산업경영공학과','가군',5,6.2,5,891.7,891.3,'N',[82,87,81,83,3,3]],
    ['환경공학과','가군',6,6.0,5,888.0,887.3,'N',[77,90,87,75,4,4]],
    ['건축학부','가군',18,4.8,20,890.8,889.3,'N',[79,85,84,88,4,2]],
    ['반도체시스템공학과','가군',33,4.7,22,906.3,904.6,'N',[86,81,98,77,3,1]],
    ['수학과','가군',2,11.0,0,895.9,895.9,'N',[86,92,56,90,2,3]],
    ['화학과','가군',5,10.8,5,889.8,889.4,'N',[97,74,66,83,3,1]],
    ['경영학과','가군',13,6.2,12,884.1,883.8,'H',[89,78,72,85,3,1]],
    ['국어교육과','가군',8,5.9,12,888.9,884.3,'H',[82,59,96,98,3,1]],
    ['교육학과','가군',13,5.8,9,882.3,881.2,'H',[82,76,93,82,3,3]],
    ['미디어커뮤니케이션학과','가군',2,11.5,0,891.2,891.2,'H',[77,90,90,93,3,1]],
    ['사회복지학과','가군',3,7.3,6,878.6,878.6,'H',[86,66,88,78,2,4]],
    ['사학과','가군',4,10.3,2,887.5,887.5,'H',[96,56,83,84,3,1]],
    ['문화콘텐츠문화경영학과','가군',7,7.4,9,881.1,880.3,'H',[85,74,81,88,3,4]],
    ['의예과','가군',40,6.1,31,991.5,990.8,'N',[99,99,99,98,2,1]],
    ['데이터사이언스학과','가군',16,5.3,21,897.6,892.3,'N',[75,90,85,83,3,3]],
    ['자유전공융합학부(인문)','가군',100,4.1,56,882.3,880.2,'FH',[92,87,81,59,3,2]],
    ['자유전공융합학부(자연)','가군',150,3.9,95,883.3,881.1,'FN',[93,87,87,61,3,5]],
    ['자연과학융합학부','가군',40,4.4,29,880.1,879.5,'FN',[80,95,72,61,3,1]],
    ['사회과학융합학부','가군',41,4.5,32,883.6,882.7,'FH',[83,63,92,92,3,1]],
    ['인문융합학부','가군',33,4.2,17,882.0,880.8,'FH',[70,63,99,95,2,2]],
    ['생명공학과','가군',15,4.5,16,899.0,896.2,'N',[82,85,81,85,2,1]],
    ['기계공학과','나군',18,4.9,21,905.8,900.4,'N',[97,74,92,81,3,3]],
    ['항공우주공학과','나군',6,6.5,3,915.6,912.6,'N',[95,92,66,85,2,2]],
    ['화공에너지공학부(화학공학전공)','나군',10,5.5,15,900.9,900.7,'N',[88,92,72,80,3,4],'화학공학과'],
    ['고분자공학과','나군',4,8.8,2,895.8,895.8,'N',[82,78,90,88,3,4]],
    ['사회인프라공학과','나군',7,5.7,5,892.2,890.4,'N',[91,90,87,68,4,2]],
    ['에너지자원공학과','나군',5,9.2,11,891.9,890.0,'N',[85,85,81,75,3,2]],
    ['화공에너지공학부(이차전지공학전공)','나군',16,4.9,9,892.8,892.0,'N',[83,85,90,71,3,1],'이차전지융합학과'],
    ['식품영양학과','나군',6,8.0,6,888.3,887.9,'N',[74,92,90,64,2,1]],
    ['파이낸스경영학과','나군',13,5.4,7,888.8,887.1,'H',[83,89,95,85,4,5]],
    ['아태물류학부','나군',6,5.5,5,902.2,901.0,'H',[83,85,97,100,4,3]],
    ['영어교육과','나군',14,4.3,13,879.4,878.7,'H',[85,63,97,73,2,2]],
    ['사회교육과','나군',8,3.9,2,891.4,888.7,'H',[86,70,92,99,4,4]],
    ['수학교육과','나군',8,7.0,9,887.4,887.3,'N',[94,81,87,64,4,1]],
    ['정치외교학과','나군',7,6.0,6,884.8,883.9,'H',[98,57,71,71,2,1]],
    ['경제학과','나군',6,7.8,6,887.2,886.1,'H',[86,54,97,94,2,1]],
    ['한국어문학과','나군',5,8.0,5,882.4,878.9,'H',[91,59,80,83,2,1]],
    ['철학과','나군',4,8.8,3,879.7,879.7,'H',[85,74,97,68,3,2]],
    ['중국학과','나군',8,7.6,10,879.7,879.4,'H',[85,70,95,74,3,1]],
    ['간호학과','나군',42,3.3,15,898.7,894.9,'N',[97,85,66,71,2,4]],
    ['컴퓨터공학과','나군',55,3.6,31,899.8,898.2,'N',[75,96,85,83,4,5]],
    ['공학융합학부','나군',131,4.3,90,885.1,883.9,'FN',[88,95,81,59,4,2]],
    ['경영융합학부','나군',44,5.0,25,888.1,885.5,'FH',[77,85,97,74,3,3]],
    ['첨단바이오의약학과','나군',10,5.2,4,906.4,904.4,'N',[82,87,94,81,3,2]],
    ['신소재공학과','다군',9,9.4,11,904.9,900.7,'N',[75,90,87,88,3,5]],
    ['공간정보공학과','다군',13,7.2,27,892.0,889.2,'N',[80,76,90,85,3,3]],
    ['전기전자공학부','다군',20,7.5,26,914.7,912.3,'N',[83,95,87,77,3,3]],
    ['통계학과','다군',3,15.0,12,888.8,879.2,'N',[70,95,99,92,6,3]],
    ['물리학과','다군',5,15.4,15,896.8,895.9,'N',[94,90,72,66,3,2]],
    ['해양과학과','다군',5,10.8,13,886.9,886.7,'N',[91,78,75,75,2,2]],
    ['국제통상학과','다군',7,13.3,15,893.8,891.5,'H',[89,57,93,95,2,3]],
    ['행정학과','다군',7,11.3,17,893.8,886.0,'H',[98,78,90,55,4,2]],
    ['소비자학과','다군',8,10.5,21,880.9,880.8,'H',[80,74,87,90,2,3]],
    ['아동심리학과','다군',7,18.0,13,883.4,883.1,'H',[82,59,97,93,2,2]],
    ['일본언어문화학과','다군',7,11.6,9,879.2,878.9,'H',[65,78,98,98,2,1]],
    ['의류디자인학과(일반)','다군',8,10.8,10,887.1,885.9,'H',[88,74,80,85,2,4]],
    ['인공지능공학과','다군',32,7.4,63,900.4,895.2,'N',[77,90,84,88,4,3]],
    ['스마트모빌리티공학과','다군',13,9.8,21,898.8,897.4,'N',[85,95,85,83,5,2]],
    ['생명과학과','다군',12,8.8,27,894.8,891.6,'N',[83,85,87,71,2,1]]
  ];

  CALCULATION_RULES['inha-regular-2025']={
    id:'inha-regular-2025',university:'inha',admissionType:'regular',appliesToYears:[2025],status:'verified',
    engine:'inha-profile-percentile',
    scoreScale:1000,
    reflectedSubjects:['국어','수학','탐구1','탐구2','영어','한국사'],
    weightOptionsByTrack:WEIGHT_OPTIONS_BY_TRACK,
    notes:['학생 성적 대비 실제 계산은 인하대 전용 백분위 환산 엔진(inhaReferenceScore/inhaProfileReferenceScore)을 그대로 사용합니다.','이 레코드는 계열별 반영비율 구조를 문서화하는 참조용입니다.'],
    sources:[{org:'대학어디가',title:'2025학년도 정시 입시결과',url:INHA_SOURCE_URL,publishedDate:'2025-05-01',checkedDate:'2026-09-01',page:null}]
  };

  // 수시(학생부교과 지역균형·학생부종합 인하미래인재) 2024·2025학년도 결과 — 대학어디가 "평가기준 및 입시결과" 탭에서 확인.
  // avgGradeReference는 대학이 공개한 "최종등록자 교과성적 학생부등급"(원 석차등급 평균, 등급→점수 환산 이전) 그대로이며,
  // 학생 반영교과 원점수 평균과 같은 척도라 시뮬레이터 비교에 바로 쓸 수 있다(가천대와 동일한 방식).
  // 대학환산점수(1,000점 만점)는 참고용으로만 표시하고 계산에는 쓰지 않는다(척도가 달라 직접 비교 불가 — 가천대 사례와 동일 이유).
  const JIGYEOK_BY_NAME={"기계공학과":{"2024":{"recruit":35,"competition":5.4,"prelim":36,"rawGrade50":2.4,"rawGrade70":2.49,"uniCut50":997.47,"uniCut70":997.33},"2025":{"recruit":32,"competition":7.1,"prelim":43,"rawGrade50":2.29,"rawGrade70":2.36,"uniCut50_25rule":997.37,"uniCut70_25rule":997.25,"uniCut50_26rule":973.66,"uniCut70_26rule":972.53}},"항공우주공학과":{"2024":{"recruit":12,"competition":7.8,"prelim":24,"rawGrade50":2.53,"rawGrade70":2.55,"uniCut50":997.23,"uniCut70":997.16},"2025":{"recruit":12,"competition":9.7,"prelim":19,"rawGrade50":2.23,"rawGrade70":2.3,"uniCut50_25rule":997.45,"uniCut70_25rule":997.4,"uniCut50_26rule":974.53,"uniCut70_26rule":973.98}},"조선해양공학과":{"2024":{"recruit":12,"competition":9.4,"prelim":13,"rawGrade50":2.81,"rawGrade70":2.85,"uniCut50":996.68,"uniCut70":996.59},"2025":{"recruit":12,"competition":11.6,"prelim":11,"rawGrade50":2.62,"rawGrade70":2.63,"uniCut50_25rule":996.71,"uniCut70_25rule":996.67,"uniCut50_26rule":967.08,"uniCut70_26rule":966.67}},"산업경영공학과":{"2024":{"recruit":10,"competition":4.3,"prelim":6,"rawGrade50":2.44,"rawGrade70":2.57,"uniCut50":997.32,"uniCut70":997.08},"2025":{"recruit":10,"competition":7.8,"prelim":11,"rawGrade50":2.43,"rawGrade70":2.46,"uniCut50_25rule":997.13,"uniCut70_25rule":997.05,"uniCut50_26rule":971.34,"uniCut70_26rule":970.53}},"화공에너지공학부(화학공학전공)":{"2024":{"recruit":21,"competition":3.7,"prelim":17,"rawGrade50":1.96,"rawGrade70":2.03,"uniCut50":998.3,"uniCut70":998.06},"2025":{"recruit":21,"competition":7.6,"prelim":29,"rawGrade50":1.89,"rawGrade70":2,"uniCut50_25rule":998.21,"uniCut70_25rule":998,"uniCut50_26rule":982.12,"uniCut70_26rule":980}},"고분자공학과":{"2024":{"recruit":9,"competition":4,"prelim":7,"rawGrade50":2.28,"rawGrade70":2.33,"uniCut50":997.67,"uniCut70":997.6},"2025":{"recruit":9,"competition":7.6,"prelim":13,"rawGrade50":2.32,"rawGrade70":2.4,"uniCut50_25rule":997.37,"uniCut70_25rule":997.2,"uniCut50_26rule":973.68,"uniCut70_26rule":971.95}},"신소재공학과":{"2024":{"recruit":18,"competition":3.9,"prelim":21,"rawGrade50":2.2,"rawGrade70":2.31,"uniCut50":997.81,"uniCut70":997.68},"2025":{"recruit":19,"competition":7.6,"prelim":23,"rawGrade50":2.09,"rawGrade70":2.21,"uniCut50_25rule":997.82,"uniCut70_25rule":997.59,"uniCut50_26rule":978.16,"uniCut70_26rule":975.87}},"사회인프라공학과":{"2024":{"recruit":13,"competition":7.2,"prelim":19,"rawGrade50":2.75,"rawGrade70":2.78,"uniCut50":996.86,"uniCut70":996.74},"2025":{"recruit":13,"competition":10.8,"prelim":21,"rawGrade50":2.48,"rawGrade70":2.58,"uniCut50_25rule":996.97,"uniCut70_25rule":996.85,"uniCut50_26rule":969.68,"uniCut70_26rule":968.5}},"환경공학과":{"2024":{"recruit":9,"competition":4.8,"prelim":13,"rawGrade50":2.34,"rawGrade70":2.46,"uniCut50":997.59,"uniCut70":997.46},"2025":{"recruit":9,"competition":7.3,"prelim":12,"rawGrade50":2.19,"rawGrade70":2.31,"uniCut50_25rule":997.49,"uniCut70_25rule":997.34,"uniCut50_26rule":974.9,"uniCut70_26rule":973.4}},"공간정보공학과":{"2024":{"recruit":7,"competition":8,"prelim":5,"rawGrade50":2.62,"rawGrade70":2.68,"uniCut50":997.05,"uniCut70":996.95},"2025":{"recruit":7,"competition":9.4,"prelim":3,"rawGrade50":2.49,"rawGrade70":2.52,"uniCut50_25rule":996.94,"uniCut70_25rule":996.93,"uniCut50_26rule":969.43,"uniCut70_26rule":969.34}},"건축학부":{"2024":{"recruit":16,"competition":6.9,"prelim":23,"rawGrade50":2.67,"rawGrade70":2.79,"uniCut50":996.85,"uniCut70":996.76},"2025":{"recruit":16,"competition":10.1,"prelim":11,"rawGrade50":2.48,"rawGrade70":2.49,"uniCut50_25rule":997.05,"uniCut70_25rule":997.01,"uniCut50_26rule":970.47,"uniCut70_26rule":970.11}},"에너지자원공학과":{"2024":{"recruit":5,"competition":8.8,"prelim":3,"rawGrade50":2.34,"rawGrade70":2.39,"uniCut50":997.56,"uniCut70":997.52},"2025":{"recruit":5,"competition":8.6,"prelim":4,"rawGrade50":2.28,"rawGrade70":2.38,"uniCut50_25rule":997.25,"uniCut70_25rule":997.13,"uniCut50_26rule":972.5,"uniCut70_26rule":971.25}},"전기전자공학부":{"2025":{"recruit":53,"competition":9.5,"prelim":76,"rawGrade50":2.19,"rawGrade70":2.32,"uniCut50_25rule":997.63,"uniCut70_25rule":997.33,"uniCut50_26rule":976.3,"uniCut70_26rule":973.26}},"반도체시스템공학과":{"2024":{"recruit":8,"competition":5.1,"prelim":4,"rawGrade50":2.62,"rawGrade70":2.65,"uniCut50":997.08,"uniCut70":997.05},"2025":{"recruit":11,"competition":14.7,"prelim":8,"rawGrade50":2.16,"rawGrade70":2.26,"uniCut50_25rule":997.65,"uniCut70_25rule":997.49,"uniCut50_26rule":976.52,"uniCut70_26rule":974.89}},"화공에너지공학부(이차전지공학전공)":{"2025":{"recruit":5,"competition":12,"prelim":4,"rawGrade50":2.48,"rawGrade70":2.69,"uniCut50_25rule":997.05,"uniCut70_25rule":996.63,"uniCut50_26rule":970.45,"uniCut70_26rule":966.26}},"수학과":{"2024":{"recruit":7,"competition":11.3,"prelim":14,"rawGrade50":2.58,"rawGrade70":2.66,"uniCut50":997.02,"uniCut70":996.99},"2025":{"recruit":7,"competition":8.6,"prelim":15,"rawGrade50":2.43,"rawGrade70":2.48,"uniCut50_25rule":997.14,"uniCut70_25rule":996.91,"uniCut50_26rule":971.37,"uniCut70_26rule":969.07}},"통계학과":{"2024":{"recruit":6,"competition":6.3,"prelim":5,"rawGrade50":2.86,"rawGrade70":3.1,"uniCut50":996.62,"uniCut70":996.07},"2025":{"recruit":6,"competition":18,"prelim":6,"rawGrade50":2.53,"rawGrade70":2.6,"uniCut50_25rule":996.94,"uniCut70_25rule":996.76,"uniCut50_26rule":969.44,"uniCut70_26rule":967.6}},"물리학과":{"2024":{"recruit":8,"competition":6.8,"prelim":12,"rawGrade50":2.77,"rawGrade70":2.97,"uniCut50":996.79,"uniCut70":996.22},"2025":{"recruit":8,"competition":13.5,"prelim":6,"rawGrade50":2.39,"rawGrade70":2.45,"uniCut50_25rule":997.23,"uniCut70_25rule":997.1,"uniCut50_26rule":972.29,"uniCut70_26rule":971.03}},"화학과":{"2024":{"recruit":10,"competition":4.6,"prelim":8,"rawGrade50":2.34,"rawGrade70":2.47,"uniCut50":997.63,"uniCut70":997.33},"2025":{"recruit":10,"competition":9.2,"prelim":13,"rawGrade50":2.21,"rawGrade70":2.28,"uniCut50_25rule":997.57,"uniCut70_25rule":997.35,"uniCut50_26rule":975.71,"uniCut70_26rule":973.54}},"해양과학과":{"2024":{"recruit":7,"competition":4.6,"prelim":4,"rawGrade50":2.49,"rawGrade70":2.53,"uniCut50":997.26,"uniCut70":997.26},"2025":{"recruit":7,"competition":6.9,"prelim":5,"rawGrade50":2.24,"rawGrade70":2.34,"uniCut50_25rule":997.43,"uniCut70_25rule":997.33,"uniCut50_26rule":974.29,"uniCut70_26rule":973.26}},"식품영양학과":{"2024":{"recruit":9,"competition":7,"prelim":7,"rawGrade50":2.48,"rawGrade70":2.53,"uniCut50":997.32,"uniCut70":997.23},"2025":{"recruit":9,"competition":7.2,"prelim":7,"rawGrade50":2.59,"rawGrade70":2.67,"uniCut50_25rule":996.78,"uniCut70_25rule":996.52,"uniCut50_26rule":967.79,"uniCut70_26rule":965.21}},"경영학과":{"2024":{"recruit":35,"competition":5.4,"prelim":64,"rawGrade50":2.47,"rawGrade70":2.52,"uniCut50":997.29,"uniCut70":997.18},"2025":{"recruit":30,"competition":7.3,"prelim":74,"rawGrade50":2.29,"rawGrade70":2.34,"uniCut50_25rule":997.39,"uniCut70_25rule":997.33,"uniCut50_26rule":973.88,"uniCut70_26rule":973.27}},"파이낸스경영학과":{"2025":{"recruit":7,"competition":9.3,"prelim":11,"rawGrade50":2.56,"rawGrade70":2.59,"uniCut50_25rule":996.8,"uniCut70_25rule":996.75,"uniCut50_26rule":968.04,"uniCut70_26rule":967.52}},"아태물류학부":{"2024":{"recruit":15,"competition":4.9,"prelim":42,"rawGrade50":2.31,"rawGrade70":2.4,"uniCut50":997.64,"uniCut70":997.45},"2025":{"recruit":15,"competition":6.5,"prelim":31,"rawGrade50":2.13,"rawGrade70":2.19,"uniCut50_25rule":997.75,"uniCut70_25rule":997.61,"uniCut50_26rule":977.52,"uniCut70_26rule":976.09}},"국제통상학과":{"2024":{"recruit":14,"competition":6.6,"prelim":24,"rawGrade50":2.49,"rawGrade70":2.56,"uniCut50":997.33,"uniCut70":997.06},"2025":{"recruit":14,"competition":8.8,"prelim":18,"rawGrade50":2.4,"rawGrade70":2.49,"uniCut50_25rule":997.16,"uniCut70_25rule":997.02,"uniCut50_26rule":971.58,"uniCut70_26rule":970.22}},"국어교육과":{"2024":{"recruit":5,"competition":7.2,"prelim":19,"rawGrade50":2.3,"rawGrade70":2.3,"uniCut50":997.47,"uniCut70":997.47},"2025":{"recruit":5,"competition":8.8,"prelim":21,"rawGrade50":1.88,"rawGrade70":2.45,"uniCut50_25rule":998.24,"uniCut70_25rule":997.09,"uniCut50_26rule":982.38,"uniCut70_26rule":970.91}},"영어교육과":{"2024":{"recruit":5,"competition":6.6,"prelim":10,"rawGrade50":2.26,"rawGrade70":3.29,"uniCut50":997.7,"uniCut70":995.73},"2025":{"recruit":5,"competition":18,"prelim":15,"rawGrade50":1.98,"rawGrade70":1.98,"uniCut50_25rule":997.95,"uniCut70_25rule":997.95,"uniCut50_26rule":979.52,"uniCut70_26rule":979.52}},"사회교육과":{"2024":{"recruit":5,"competition":5.6,"prelim":12,"rawGrade50":2.26,"rawGrade70":2.29,"uniCut50":997.59,"uniCut70":997.59},"2025":{"recruit":5,"competition":7.2,"prelim":11,"rawGrade50":2.17,"rawGrade70":2.21,"uniCut50_25rule":997.67,"uniCut70_25rule":997.58,"uniCut50_26rule":976.67,"uniCut70_26rule":975.79}},"교육학과":{"2024":{"recruit":5,"competition":6.2,"prelim":9,"rawGrade50":2.29,"rawGrade70":2.33,"uniCut50":997.65,"uniCut70":997.41},"2025":{"recruit":5,"competition":6.2,"prelim":9,"rawGrade50":2.06,"rawGrade70":2.48,"uniCut50_25rule":997.89,"uniCut70_25rule":996.95,"uniCut50_26rule":978.85,"uniCut70_26rule":969.48}},"수학교육과":{"2024":{"recruit":5,"competition":5.2,"prelim":7,"rawGrade50":2.2,"rawGrade70":2.2,"uniCut50":997.82,"uniCut70":997.78},"2025":{"recruit":5,"competition":6.2,"prelim":5,"rawGrade50":2.13,"rawGrade70":2.18,"uniCut50_25rule":997.73,"uniCut70_25rule":997.64,"uniCut50_26rule":977.33,"uniCut70_26rule":976.43}},"행정학과":{"2024":{"recruit":13,"competition":5.2,"prelim":25,"rawGrade50":2.45,"rawGrade70":2.82,"uniCut50":997.21,"uniCut70":996.43},"2025":{"recruit":13,"competition":9.8,"prelim":23,"rawGrade50":2.47,"rawGrade70":2.54,"uniCut50_25rule":996.98,"uniCut70_25rule":996.9,"uniCut50_26rule":969.75,"uniCut70_26rule":969.03}},"정치외교학과":{"2024":{"recruit":10,"competition":5.5,"prelim":25,"rawGrade50":2.53,"rawGrade70":2.74,"uniCut50":996.95,"uniCut70":996.72},"2025":{"recruit":10,"competition":8,"prelim":21,"rawGrade50":2.39,"rawGrade70":2.45,"uniCut50_25rule":997.21,"uniCut70_25rule":997.02,"uniCut50_26rule":972.12,"uniCut70_26rule":970.2}},"미디어커뮤니케이션학과":{"2024":{"recruit":10,"competition":6.4,"prelim":26,"rawGrade50":2.4,"rawGrade70":2.46,"uniCut50":997.44,"uniCut70":997.25},"2025":{"recruit":10,"competition":6.6,"prelim":22,"rawGrade50":2.25,"rawGrade70":2.48,"uniCut50_25rule":997.41,"uniCut70_25rule":996.96,"uniCut50_26rule":974.12,"uniCut70_26rule":969.59}},"경제학과":{"2024":{"recruit":13,"competition":6.2,"prelim":30,"rawGrade50":2.45,"rawGrade70":2.49,"uniCut50":997.34,"uniCut70":997.32},"2025":{"recruit":13,"competition":7.6,"prelim":26,"rawGrade50":2.44,"rawGrade70":2.53,"uniCut50_25rule":997.13,"uniCut70_25rule":996.94,"uniCut50_26rule":971.29,"uniCut70_26rule":969.39}},"소비자학과":{"2024":{"recruit":5,"competition":24.8,"prelim":1,"rawGrade50":2.41,"rawGrade70":2.48,"uniCut50":997.37,"uniCut70":997.24},"2025":{"recruit":5,"competition":13,"prelim":7,"rawGrade50":2.71,"rawGrade70":2.79,"uniCut50_25rule":996.35,"uniCut70_25rule":996.33,"uniCut50_26rule":963.46,"uniCut70_26rule":963.27}},"아동심리학과":{"2024":{"recruit":5,"competition":5.2,"prelim":11,"rawGrade50":2.55,"rawGrade70":2.55,"uniCut50":997.08,"uniCut70":997.08},"2025":{"recruit":5,"competition":6.6,"prelim":12,"rawGrade50":2.46,"rawGrade70":2.47,"uniCut50_25rule":997.08,"uniCut70_25rule":997.06,"uniCut50_26rule":970.77,"uniCut70_26rule":970.63}},"사회복지학과":{"2024":{"recruit":5,"competition":4.8,"prelim":7,"rawGrade50":2.48,"rawGrade70":2.68,"uniCut50":997.31,"uniCut70":996.83},"2025":{"recruit":5,"competition":7.8,"prelim":9,"rawGrade50":2.77,"rawGrade70":2.77,"uniCut50_25rule":996.41,"uniCut70_25rule":996.41,"uniCut50_26rule":964.1,"uniCut70_26rule":964.1}},"한국어문학과":{"2024":{"recruit":8,"competition":9.6,"prelim":10,"rawGrade50":2.67,"rawGrade70":2.76,"uniCut50":996.88,"uniCut70":996.66},"2025":{"recruit":8,"competition":8.3,"prelim":10,"rawGrade50":2.66,"rawGrade70":2.7,"uniCut50_25rule":996.63,"uniCut70_25rule":996.5,"uniCut50_26rule":966.32,"uniCut70_26rule":965}},"사학과":{"2024":{"recruit":6,"competition":10.7,"prelim":11,"rawGrade50":2.51,"rawGrade70":2.66,"uniCut50":997.19,"uniCut70":996.96},"2025":{"recruit":6,"competition":5.3,"prelim":12,"rawGrade50":2.49,"rawGrade70":2.55,"uniCut50_25rule":996.96,"uniCut70_25rule":996.76,"uniCut50_26rule":969.6,"uniCut70_26rule":967.6}},"철학과":{"2024":{"recruit":6,"competition":7.7,"prelim":5,"rawGrade50":2.74,"rawGrade70":3.06,"uniCut50":996.78,"uniCut70":996.22},"2025":{"recruit":6,"competition":9.3,"prelim":7,"rawGrade50":2.59,"rawGrade70":2.61,"uniCut50_25rule":996.76,"uniCut70_25rule":996.68,"uniCut50_26rule":967.63,"uniCut70_26rule":966.75}},"중국학과":{"2024":{"recruit":9,"competition":12.9,"prelim":20,"rawGrade50":2.82,"rawGrade70":3,"uniCut50":996.39,"uniCut70":996.29},"2025":{"recruit":10,"competition":10.5,"prelim":6,"rawGrade50":2.59,"rawGrade70":2.66,"uniCut50_25rule":996.83,"uniCut70_25rule":996.68,"uniCut50_26rule":968.25,"uniCut70_26rule":966.81}},"일본언어문화학과":{"2024":{"recruit":9,"competition":11.9,"prelim":4,"rawGrade50":2.73,"rawGrade70":2.8,"uniCut50":996.65,"uniCut70":996.6},"2025":{"recruit":9,"competition":8.4,"prelim":5,"rawGrade50":2.57,"rawGrade70":2.62,"uniCut50_25rule":996.76,"uniCut70_25rule":996.71,"uniCut50_26rule":967.55,"uniCut70_26rule":967.07}},"영미유럽인문융합학부":{"2025":{"recruit":15,"competition":6.9,"prelim":22,"rawGrade50":2.52,"rawGrade70":2.62,"uniCut50_25rule":996.95,"uniCut70_25rule":996.76,"uniCut50_26rule":969.54,"uniCut70_26rule":967.6}},"문화콘텐츠문화경영학과":{"2024":{"recruit":12,"competition":6.1,"prelim":28,"rawGrade50":2.56,"rawGrade70":2.63,"uniCut50":997.04,"uniCut70":996.95},"2025":{"recruit":13,"competition":7.2,"prelim":34,"rawGrade50":2.4,"rawGrade70":2.51,"uniCut50_25rule":997.17,"uniCut70_25rule":996.98,"uniCut50_26rule":971.72,"uniCut70_26rule":969.75}},"의예과":{"2024":{"recruit":9,"competition":9.4,"prelim":12,"rawGrade50":1.08,"rawGrade70":1.1,"uniCut50":999.85,"uniCut70":999.82},"2025":{"recruit":26,"competition":13.6,"prelim":24,"rawGrade50":1.1,"rawGrade70":1.13,"uniCut50_25rule":999.81,"uniCut70_25rule":999.73,"uniCut50_26rule":998.06,"uniCut70_26rule":997.33}},"간호학과":{"2024":{"recruit":17,"competition":3.6,"prelim":14,"rawGrade50":2.1,"rawGrade70":2.16,"uniCut50":997.98,"uniCut70":997.8},"2025":{"recruit":17,"competition":14.4,"prelim":30,"rawGrade50":1.97,"rawGrade70":2,"uniCut50_25rule":998.06,"uniCut70_25rule":998,"uniCut50_26rule":980.62,"uniCut70_26rule":980}},"의류디자인학과(일반)":{"2024":{"recruit":6,"competition":10,"prelim":11,"rawGrade50":2.64,"rawGrade70":2.73,"uniCut50":996.83,"uniCut70":996.77},"2025":{"recruit":6,"competition":8,"prelim":13,"rawGrade50":2.51,"rawGrade70":2.58,"uniCut50_25rule":996.99,"uniCut70_25rule":996.84,"uniCut50_26rule":969.88,"uniCut70_26rule":968.43}},"인공지능공학과":{"2024":{"recruit":9,"competition":6.1,"prelim":15,"rawGrade50":2.4,"rawGrade70":2.42,"uniCut50":997.47,"uniCut70":997.43},"2025":{"recruit":11,"competition":7.4,"prelim":8,"rawGrade50":2.18,"rawGrade70":2.29,"uniCut50_25rule":997.59,"uniCut70_25rule":997.41,"uniCut50_26rule":975.88,"uniCut70_26rule":974.13}},"데이터사이언스학과":{"2024":{"recruit":10,"competition":9.6,"prelim":12,"rawGrade50":2.4,"rawGrade70":2.47,"uniCut50":997.43,"uniCut70":997.28},"2025":{"recruit":10,"competition":9.7,"prelim":7,"rawGrade50":2.31,"rawGrade70":2.33,"uniCut50_25rule":997.37,"uniCut70_25rule":997.34,"uniCut50_26rule":973.72,"uniCut70_26rule":973.4}},"스마트모빌리티공학과":{"2024":{"recruit":7,"competition":11.6,"prelim":18,"rawGrade50":2.55,"rawGrade70":2.57,"uniCut50":997.2,"uniCut70":997.18},"2025":{"recruit":7,"competition":8.7,"prelim":11,"rawGrade50":2.43,"rawGrade70":2.44,"uniCut50_25rule":997.11,"uniCut70_25rule":997.06,"uniCut50_26rule":971.11,"uniCut70_26rule":970.63}},"컴퓨터공학과":{"2024":{"recruit":38,"competition":3.7,"prelim":28,"rawGrade50":2.26,"rawGrade70":2.38,"uniCut50":997.69,"uniCut70":997.49},"2025":{"recruit":38,"competition":6.2,"prelim":45,"rawGrade50":2.29,"rawGrade70":2.44,"uniCut50_25rule":997.42,"uniCut70_25rule":997.05,"uniCut50_26rule":974.23,"uniCut70_26rule":970.53}},"자유전공융합학부":{"2025":{"recruit":20,"competition":11.8,"prelim":30,"rawGrade50":2.16,"rawGrade70":2.21,"uniCut50_25rule":997.69,"uniCut70_25rule":997.58,"uniCut50_26rule":976.86,"uniCut70_26rule":975.8}},"생명공학과":{"2024":{"recruit":9,"competition":4.9,"prelim":10,"rawGrade50":1.92,"rawGrade70":2.29,"uniCut50":998.35,"uniCut70":997.67},"2025":{"recruit":8,"competition":11,"prelim":8,"rawGrade50":1.87,"rawGrade70":1.92,"uniCut50_25rule":998.27,"uniCut70_25rule":998.16,"uniCut50_26rule":982.67,"uniCut70_26rule":981.59}},"생명과학과":{"2024":{"recruit":6,"competition":7.2,"prelim":7,"rawGrade50":2.07,"rawGrade70":2.08,"uniCut50":998.07,"uniCut70":998.06},"2025":{"recruit":5,"competition":7.4,"prelim":8,"rawGrade50":2.09,"rawGrade70":2.09,"uniCut50_25rule":997.82,"uniCut70_25rule":997.82,"uniCut50_26rule":978.22,"uniCut70_26rule":978.22}},"전기공학과":{"2024":{"recruit":14,"competition":4.4,"prelim":18,"rawGrade50":2.35,"rawGrade70":2.41,"uniCut50":997.65,"uniCut70":997.46}},"전자공학과":{"2024":{"recruit":15,"competition":4,"prelim":20,"rawGrade50":2.04,"rawGrade70":2.05,"uniCut50":998.13,"uniCut70":998.11}},"정보통신공학과":{"2024":{"recruit":18,"competition":4.3,"prelim":17,"rawGrade50":2.46,"rawGrade70":2.74,"uniCut50":997.37,"uniCut70":996.76}},"글로벌금융학과":{"2024":{"recruit":7,"competition":13,"prelim":15,"rawGrade50":2.59,"rawGrade70":2.59,"uniCut50":997.1,"uniCut70":997.09}},"영어영문학과":{"2024":{"recruit":11,"competition":4.7,"prelim":11,"rawGrade50":2.63,"rawGrade70":2.64,"uniCut50":996.97,"uniCut70":996.96}},"프랑스언어문화학과":{"2024":{"recruit":6,"competition":9.3,"prelim":4,"rawGrade50":2.76,"rawGrade70":2.82,"uniCut50":996.73,"uniCut70":996.67}}};
  const JONGHAP_BY_NAME={"기계공학과":{"2024":{"recruit":40,"competition":11.1,"prelim":29,"rawGrade50":2.81,"rawGrade70":3.03},"2025":{"recruit":40,"competition":11.7,"prelim":34,"rawGrade50":2.73,"rawGrade70":2.86}},"항공우주공학과":{"2024":{"recruit":19,"competition":15.7,"prelim":11,"rawGrade50":2.58,"rawGrade70":2.69},"2025":{"recruit":19,"competition":13.5,"prelim":14,"rawGrade50":2.59,"rawGrade70":2.74}},"조선해양공학과":{"2024":{"recruit":19,"competition":6.8,"prelim":4,"rawGrade50":3.72,"rawGrade70":3.8},"2025":{"recruit":19,"competition":10.4,"prelim":4,"rawGrade50":3.43,"rawGrade70":3.57}},"산업경영공학과":{"2024":{"recruit":15,"competition":8.9,"prelim":15,"rawGrade50":3.07,"rawGrade70":3.22},"2025":{"recruit":15,"competition":10.4,"prelim":6,"rawGrade50":3,"rawGrade70":3.29}},"화공에너지공학부(화학공학전공)":{"2024":{"recruit":34,"competition":9.2,"prelim":20,"rawGrade50":2.33,"rawGrade70":2.49},"2025":{"recruit":33,"competition":8.6,"prelim":24,"rawGrade50":2.29,"rawGrade70":2.44}},"고분자공학과":{"2024":{"recruit":15,"competition":11.4,"prelim":6,"rawGrade50":2.82,"rawGrade70":2.9},"2025":{"recruit":15,"competition":10.8,"prelim":2,"rawGrade50":2.66,"rawGrade70":2.73}},"신소재공학과":{"2024":{"recruit":32,"competition":10.3,"prelim":12,"rawGrade50":2.66,"rawGrade70":2.73},"2025":{"recruit":32,"competition":10.1,"prelim":27,"rawGrade50":2.57,"rawGrade70":2.7}},"사회인프라공학과":{"2024":{"recruit":23,"competition":11.3,"prelim":15,"rawGrade50":3.14,"rawGrade70":3.31},"2025":{"recruit":23,"competition":10.8,"prelim":20,"rawGrade50":3.38,"rawGrade70":3.63}},"환경공학과":{"2024":{"recruit":14,"competition":19.6,"prelim":17,"rawGrade50":2.69,"rawGrade70":2.78},"2025":{"recruit":14,"competition":15.4,"prelim":4,"rawGrade50":2.62,"rawGrade70":2.71}},"공간정보공학과":{"2024":{"recruit":13,"competition":17.6,"prelim":2,"rawGrade50":3.25,"rawGrade70":3.42},"2025":{"recruit":12,"competition":10.2,"prelim":4,"rawGrade50":3.44,"rawGrade70":3.73}},"건축학부":{"2024":{"recruit":24,"competition":14.8,"prelim":25,"rawGrade50":2.97,"rawGrade70":3.07},"2025":{"recruit":24,"competition":13.5,"prelim":22,"rawGrade50":2.88,"rawGrade70":2.94}},"에너지자원공학과":{"2024":{"recruit":8,"competition":9.6,"prelim":3,"rawGrade50":2.51,"rawGrade70":2.72},"2025":{"recruit":8,"competition":8.9,"prelim":4,"rawGrade50":2.88,"rawGrade70":2.91}},"전기전자공학부":{"2025":{"recruit":71,"competition":9.6,"prelim":57,"rawGrade50":2.55,"rawGrade70":2.68}},"반도체시스템공학과":{"2024":{"recruit":13,"competition":9.8,"prelim":4,"rawGrade50":2.87,"rawGrade70":3.09},"2025":{"recruit":23,"competition":9.3,"prelim":8,"rawGrade50":2.78,"rawGrade70":2.91}},"화공에너지공학부(이차전지공학전공)":{"2025":{"recruit":13,"competition":9.5,"prelim":12,"rawGrade50":2.98,"rawGrade70":3.45}},"수학과":{"2024":{"recruit":11,"competition":14.6,"prelim":3,"rawGrade50":3.07,"rawGrade70":3.11},"2025":{"recruit":11,"competition":11.3,"prelim":10,"rawGrade50":2.8,"rawGrade70":2.94}},"통계학과":{"2024":{"recruit":10,"competition":9.9,"prelim":10,"rawGrade50":2.94,"rawGrade70":3.03},"2025":{"recruit":9,"competition":10.7,"prelim":5,"rawGrade50":3.05,"rawGrade70":3.2}},"물리학과":{"2024":{"recruit":13,"competition":10.6,"prelim":16,"rawGrade50":3.03,"rawGrade70":3.34},"2025":{"recruit":13,"competition":10.2,"prelim":17,"rawGrade50":2.99,"rawGrade70":3.17}},"화학과":{"2024":{"recruit":16,"competition":16.1,"prelim":12,"rawGrade50":2.65,"rawGrade70":2.73},"2025":{"recruit":16,"competition":12.8,"prelim":9,"rawGrade50":2.69,"rawGrade70":2.73}},"해양과학과":{"2024":{"recruit":11,"competition":11.5,"prelim":0,"rawGrade50":3.23,"rawGrade70":3.33},"2025":{"recruit":11,"competition":15.1,"prelim":2,"rawGrade50":2.73,"rawGrade70":2.83}},"식품영양학과":{"2024":{"recruit":14,"competition":16.9,"prelim":7,"rawGrade50":3.27,"rawGrade70":3.43},"2025":{"recruit":14,"competition":25.6,"prelim":7,"rawGrade50":2.86,"rawGrade70":2.89}},"경영학과":{"2024":{"recruit":40,"competition":18.3,"prelim":32,"rawGrade50":2.87,"rawGrade70":2.93},"2025":{"recruit":40,"competition":11.9,"prelim":16,"rawGrade50":2.95,"rawGrade70":3.16}},"파이낸스경영학과":{"2025":{"recruit":11,"competition":9.8,"prelim":5,"rawGrade50":3.18,"rawGrade70":3.23}},"아태물류학부":{"2024":{"recruit":24,"competition":7.6,"prelim":11,"rawGrade50":2.99,"rawGrade70":3.33},"2025":{"recruit":22,"competition":7.4,"prelim":6,"rawGrade50":3.2,"rawGrade70":3.32}},"국제통상학과":{"2024":{"recruit":23,"competition":12.2,"prelim":17,"rawGrade50":3.13,"rawGrade70":3.32},"2025":{"recruit":22,"competition":8.3,"prelim":13,"rawGrade50":3.14,"rawGrade70":3.34}},"국어교육과":{"2024":{"recruit":8,"competition":13.9,"prelim":6,"rawGrade50":2.33,"rawGrade70":2.44},"2025":{"recruit":8,"competition":7.8,"prelim":6,"rawGrade50":2.49,"rawGrade70":2.54}},"영어교육과":{"2024":{"recruit":7,"competition":13.7,"prelim":7,"rawGrade50":2.46,"rawGrade70":2.52},"2025":{"recruit":7,"competition":10.1,"prelim":3,"rawGrade50":2.64,"rawGrade70":2.68}},"사회교육과":{"2024":{"recruit":7,"competition":20.6,"prelim":6,"rawGrade50":2.73,"rawGrade70":2.75},"2025":{"recruit":7,"competition":15.7,"prelim":3,"rawGrade50":2.48,"rawGrade70":2.51}},"교육학과":{"2024":{"recruit":7,"competition":14.7,"prelim":8,"rawGrade50":2.48,"rawGrade70":2.55},"2025":{"recruit":7,"competition":11.1,"prelim":8,"rawGrade50":2.58,"rawGrade70":2.62}},"수학교육과":{"2024":{"recruit":7,"competition":16.1,"prelim":3,"rawGrade50":2.45,"rawGrade70":2.45},"2025":{"recruit":7,"competition":10.3,"prelim":4,"rawGrade50":2.53,"rawGrade70":2.62}},"행정학과":{"2024":{"recruit":20,"competition":12.4,"prelim":12,"rawGrade50":3.09,"rawGrade70":3.29},"2025":{"recruit":20,"competition":10.8,"prelim":16,"rawGrade50":2.93,"rawGrade70":3.11}},"정치외교학과":{"2024":{"recruit":14,"competition":16.7,"prelim":7,"rawGrade50":4.35,"rawGrade70":4.75},"2025":{"recruit":14,"competition":19.4,"prelim":10,"rawGrade50":2.9,"rawGrade70":3.05}},"미디어커뮤니케이션학과":{"2024":{"recruit":15,"competition":18.9,"prelim":7,"rawGrade50":2.63,"rawGrade70":2.71},"2025":{"recruit":15,"competition":13.1,"prelim":9,"rawGrade50":2.59,"rawGrade70":2.67}},"경제학과":{"2024":{"recruit":20,"competition":8.3,"prelim":9,"rawGrade50":2.78,"rawGrade70":2.88},"2025":{"recruit":20,"competition":6.8,"prelim":18,"rawGrade50":3.03,"rawGrade70":3.28}},"소비자학과":{"2024":{"recruit":7,"competition":13.3,"prelim":0,"rawGrade50":2.99,"rawGrade70":3.17},"2025":{"recruit":7,"competition":11.6,"prelim":1,"rawGrade50":2.98,"rawGrade70":3.07}},"아동심리학과":{"2024":{"recruit":7,"competition":16,"prelim":5,"rawGrade50":3.07,"rawGrade70":3.09},"2025":{"recruit":8,"competition":12.6,"prelim":6,"rawGrade50":2.92,"rawGrade70":2.99}},"사회복지학과":{"2024":{"recruit":7,"competition":16.7,"prelim":5,"rawGrade50":2.71,"rawGrade70":2.84},"2025":{"recruit":8,"competition":10.4,"prelim":7,"rawGrade50":2.86,"rawGrade70":2.87}},"한국어문학과":{"2024":{"recruit":12,"competition":9.1,"prelim":6,"rawGrade50":3.3,"rawGrade70":3.46},"2025":{"recruit":12,"competition":9.2,"prelim":9,"rawGrade50":3.3,"rawGrade70":3.34}},"사학과":{"2024":{"recruit":9,"competition":14.6,"prelim":4,"rawGrade50":2.89,"rawGrade70":3},"2025":{"recruit":9,"competition":12.7,"prelim":11,"rawGrade50":2.9,"rawGrade70":3.4}},"철학과":{"2024":{"recruit":8,"competition":10.4,"prelim":5,"rawGrade50":2.99,"rawGrade70":3.12},"2025":{"recruit":8,"competition":10,"prelim":9,"rawGrade50":2.99,"rawGrade70":3}},"중국학과":{"2024":{"recruit":16,"competition":10.6,"prelim":4,"rawGrade50":5.01,"rawGrade70":5.4},"2025":{"recruit":15,"competition":16.9,"prelim":16,"rawGrade50":3.84,"rawGrade70":4.01}},"일본언어문화학과":{"2024":{"recruit":13,"competition":13.7,"prelim":4,"rawGrade50":3.34,"rawGrade70":3.67},"2025":{"recruit":13,"competition":12.3,"prelim":10,"rawGrade50":3.84,"rawGrade70":5.09}},"영미유럽인문융합학부":{"2025":{"recruit":27,"competition":15.9,"prelim":27,"rawGrade50":4.52,"rawGrade70":4.76}},"문화콘텐츠문화경영학과":{"2024":{"recruit":19,"competition":14.6,"prelim":10,"rawGrade50":2.91,"rawGrade70":2.97},"2025":{"recruit":20,"competition":12.8,"prelim":10,"rawGrade50":2.78,"rawGrade70":2.87}},"의예과":{"2024":{"recruit":16,"competition":21,"prelim":9,"rawGrade50":1.05,"rawGrade70":1.08},"2025":{"recruit":42,"competition":16.5,"prelim":27,"rawGrade50":1.13,"rawGrade70":1.41}},"간호학과":{"2024":{"recruit":25,"competition":16,"prelim":16,"rawGrade50":2.35,"rawGrade70":2.46},"2025":{"recruit":28,"competition":15.8,"prelim":20,"rawGrade50":2.4,"rawGrade70":2.42}},"의류디자인학과(일반)":{"2024":{"recruit":8,"competition":35.4,"prelim":6,"rawGrade50":3.26,"rawGrade70":3.34},"2025":{"recruit":8,"competition":23.3,"prelim":5,"rawGrade50":3.34,"rawGrade70":3.35}},"인공지능공학과":{"2024":{"recruit":17,"competition":11.4,"prelim":16,"rawGrade50":2.55,"rawGrade70":2.68},"2025":{"recruit":22,"competition":9,"prelim":7,"rawGrade50":2.74,"rawGrade70":2.95}},"데이터사이언스학과":{"2024":{"recruit":16,"competition":10.7,"prelim":11,"rawGrade50":2.95,"rawGrade70":3.14},"2025":{"recruit":15,"competition":11.9,"prelim":13,"rawGrade50":2.93,"rawGrade70":3.31}},"스마트모빌리티공학과":{"2024":{"recruit":14,"competition":12.4,"prelim":8,"rawGrade50":2.91,"rawGrade70":3.08},"2025":{"recruit":12,"competition":9.8,"prelim":6,"rawGrade50":3.09,"rawGrade70":3.14}},"컴퓨터공학과":{"2024":{"recruit":40,"competition":10.5,"prelim":29,"rawGrade50":2.4,"rawGrade70":2.7},"2025":{"recruit":40,"competition":10.2,"prelim":31,"rawGrade50":2.46,"rawGrade70":2.55}},"생명공학과":{"2024":{"recruit":14,"competition":23.8,"prelim":18,"rawGrade50":2.3,"rawGrade70":2.35},"2025":{"recruit":14,"competition":19.6,"prelim":5,"rawGrade50":2.06,"rawGrade70":2.25}},"생명과학과":{"2024":{"recruit":11,"competition":41.9,"prelim":8,"rawGrade50":2.63,"rawGrade70":2.67},"2025":{"recruit":11,"competition":26.5,"prelim":8,"rawGrade50":2.34,"rawGrade70":2.45}},"첨단바이오의약학과":{"2025":{"recruit":11,"competition":18.9,"prelim":8,"rawGrade50":2.32,"rawGrade70":2.54}},"전기공학과":{"2024":{"recruit":27,"competition":11,"prelim":16,"rawGrade50":2.74,"rawGrade70":2.96}},"전자공학과":{"2024":{"recruit":29,"competition":8.5,"prelim":33,"rawGrade50":2.53,"rawGrade70":2.62}},"정보통신공학과":{"2024":{"recruit":34,"competition":9.5,"prelim":22,"rawGrade50":2.79,"rawGrade70":2.92}},"글로벌금융학과":{"2024":{"recruit":11,"competition":12.1,"prelim":6,"rawGrade50":3.23,"rawGrade70":3.48}},"영어영문학과":{"2024":{"recruit":18,"competition":13.4,"prelim":16,"rawGrade50":4.04,"rawGrade70":4.49}},"프랑스언어문화학과":{"2024":{"recruit":9,"competition":19.7,"prelim":7,"rawGrade50":5.4,"rawGrade70":5.61}}};

  CALCULATION_RULES['inha-schoolrecord-jiyeok-2026']={
    id:'inha-schoolrecord-jiyeok-2026',university:'inha',admissionType:'schoolRecord',appliesToYears:[2025,2026],status:'draft',
    reflectedSubjects:['국어','수학','영어','사회/과학(계열별)'],
    method:'학생부교과(지역균형): 반영교과 인문(국/수/영/사·한국사포함)·자연(국/수/영/과)·자유전공융합학부(전체), 학년-학기 가중치 없음(전 학년 통합 이수단위 가중평균), 진로선택과목은 상위 3과목만 성취도→등급 변환(A=1,B=2,C=4)하여 포함, 석차등급→점수표(2026학년도: 1=10.0~9=0, 진로선택 A/B/C=10/9.8/9.4).',
    notes:['반영교과·진로선택 처리·등급→점수표는 대학어디가 "2026학년도 전형별 주요사항"(Ⅲ-1) 원문으로 확인했다.','반영점수를 1,000점 만점으로 환산하는 마지막 배수 산출식은 이미지로만 제공되어 확인하지 못해 draft로 둔다.','2025학년도 대비 2026학년도에 기본점수 미반영으로 산출식이 변경되어(대학 공지), 두 학년도를 같은 계산규칙으로 비교하지 않는다.','2024학년도는 반영방법(요소별 반영비율 등)을 별도로 확인하지 못해 결과 수치만 참고용으로 등록했다.'],
    sources:[{org:'대학어디가(adiga.kr)',title:'2026학년도 전형별 주요사항(학생부교과)',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
  };

  function jigyeokDetail(deptId,r){
    if(!r)return{details:[]};
    const years=Object.keys(r).map(Number).sort((a,b)=>b-a).map(year=>{
      const d=r[year];
      const extra=[];
      if(year===2025){
        extra.push({label:'대학환산점수(1,000점 만점, 참고용·2025학년도 산식)',value:'50%컷 '+d.uniCut50_25rule+' · 70%컷 '+d.uniCut70_25rule});
        extra.push({label:'대학환산점수(1,000점 만점, 참고용·2026학년도 산식 재계산)',value:'50%컷 '+d.uniCut50_26rule+' · 70%컷 '+d.uniCut70_26rule});
      }else{
        extra.push({label:'대학환산점수(1,000점 만점, 참고용·'+year+'학년도 산식)',value:'50%컷 '+d.uniCut50+' · 70%컷 '+d.uniCut70});
      }
      return{
        year,calculationRuleId:year===2025?'inha-schoolrecord-jiyeok-2026':null,
        recruitCount:d.recruit,competitionRatio:d.competition,additionalPass:{count:d.prelim,round:null},
        admissionMethod:'학생부교과 100%('+year+'학년도 결과. 산식이 매년 달라질 수 있어 원 등급 컷만 시뮬레이터 비교에 사용함)',csatMinimum:null,
        indicators:{avgGradeReference50:d.rawGrade50,avgGradeReference70:d.rawGrade70,extra},
        sources:[{org:'대학어디가',title:year+'학년도 학생부교과(지역균형) 전형 결과',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
      };
    });
    return{details:[{id:deptId+'-schoolrecord-jiyeok',name:'학생부교과(지역균형)',group:null,years}]};
  }
  function jonghapDetail(deptId,r){
    if(!r)return{details:[]};
    const years=Object.keys(r).map(Number).sort((a,b)=>b-a).map(year=>{
      const d=r[year];
      return{
        year,calculationRuleId:null,
        recruitCount:d.recruit,competitionRatio:d.competition,additionalPass:{count:d.prelim,round:null},
        admissionMethod:'면접형: 1단계 서류 100%(3.5배수, 의예과 3배수) → 2단계 1단계 70%+면접 30% / 서류형: 서류 100% 일괄(정성평가, 정량화 없음)',csatMinimum:null,
        indicators:{avgGradeReference50:d.rawGrade50,avgGradeReference70:d.rawGrade70,
          extra:[{label:'평가요소',value:'기초학업역량·진로탐구역량·공동체역량(면접형 30/50/20%, 서류형 50/30/20%)'}]},
        sources:[{org:'대학어디가',title:year+'학년도 학생부종합(인하미래인재) 전형 결과',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
      };
    });
    return{details:[{id:deptId+'-comprehensive-mirae',name:'학생부종합(인하미래인재)',group:null,years}]};
  }

  const departments=INHA_2025_RESULT_ROWS.map((row,index)=>{
    const [department,group,recruit,competition,extra,cut50,cut70,track,profile70,sourceDepartment]=row;
    const deptId='inha-'+String(index+1).padStart(2,'0');
    const detailId=deptId+'-regular-general';
    const jigyeok=JIGYEOK_BY_NAME[department],jonghap=JONGHAP_BY_NAME[department];
    return{
      id:deptId,name:department,college:null,
      simulationBinding:{regular:detailId,schoolRecord:jigyeok?deptId+'-schoolrecord-jiyeok':null,comprehensive:jonghap?deptId+'-comprehensive-mirae':null,essay:null},
      admissionTypes:{
        regular:{details:[{
          id:detailId,name:'일반전형(수능)',group,
          years:[{
            year:2025,calculationRuleId:'inha-regular-2025',
            recruitCount:recruit,competitionRatio:competition,additionalPass:{count:extra,round:null},
            admissionMethod:'수능 100%',csatMinimum:null,
            indicators:{universityScoreCut50:cut50,universityScoreCut70:cut70,universityScoreScale:1000,
              track,profile70,sourceDepartment:sourceDepartment||department,
              percentileCut50:null,percentileCut70:null,standardScoreCut50:null,standardScoreCut70:null,extra:[]},
            sources:[{org:'대학어디가',title:'2025학년도 정시 입시결과',url:INHA_SOURCE_URL,publishedDate:'2025-05-01',checkedDate:'2026-09-01',page:null}]
          }]
        }]},
        schoolRecord:jigyeokDetail(deptId,jigyeok),
        comprehensive:jonghapDetail(deptId,jonghap),
        essay:{details:[]}
      }
    };
  });

  // 정시 데이터에 없던 수시 전용 신설/광역 모집단위(영미유럽인문융합학부, 자유전공융합학부[계열 미분리]) 추가
  const EXTRA_DEPTS=[
    {canon:'영미유럽인문융합학부',id:'inha-61'},
    {canon:'자유전공융합학부',id:'inha-62'}
  ];
  EXTRA_DEPTS.forEach(({canon,id})=>{
    const jigyeok=JIGYEOK_BY_NAME[canon],jonghap=JONGHAP_BY_NAME[canon];
    if(!jigyeok&&!jonghap)return;
    departments.push({
      id,name:canon,college:null,
      simulationBinding:{regular:null,schoolRecord:jigyeok?id+'-schoolrecord-jiyeok':null,comprehensive:jonghap?id+'-comprehensive-mirae':null,essay:null},
      admissionTypes:{regular:{details:[]},schoolRecord:jigyeokDetail(id,jigyeok),comprehensive:jonghapDetail(id,jonghap),essay:{details:[]}}
    });
  });

  // 2024학년도에만 존재가 확인되고 2025·2026학년도 정시·수시 어느 목록에도 없는 모집단위(전공 통합·개편으로 추정되나
  // 공식 자료에 명시적 개편 표기가 없어 임의로 연결하지 않고 2024학년도 전용 별개 모집단위로 유지한다).
  const STANDALONE_2024=["전기공학과","전자공학과","정보통신공학과","글로벌금융학과","영어영문학과","프랑스언어문화학과"];
  STANDALONE_2024.forEach((canon,idx)=>{
    const jigyeok=JIGYEOK_BY_NAME[canon],jonghap=JONGHAP_BY_NAME[canon];
    if(!jigyeok&&!jonghap)return;
    const id='inha-'+String(63+idx).padStart(2,'0');
    departments.push({
      id,name:canon,college:null,
      simulationBinding:{regular:null,schoolRecord:jigyeok?id+'-schoolrecord-jiyeok':null,comprehensive:jonghap?id+'-comprehensive-mirae':null,essay:null},
      admissionTypes:{regular:{details:[]},schoolRecord:jigyeokDetail(id,jigyeok),comprehensive:jonghapDetail(id,jonghap),essay:{details:[]}}
    });
  });

  UNIVERSITY_DB.push({id:'inha',name:'인하대학교',aliases:['인하대'],region:'인천',guideUrl:INHA_2027_GUIDE_URL,departments});
})();
