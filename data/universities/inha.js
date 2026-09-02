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

  // 수시(학생부교과 지역균형·학생부종합 인하미래인재) 2025학년도 결과 — 대학어디가 "평가기준 및 입시결과" 탭에서 확인.
  // avgGradeReference는 대학이 공개한 "최종등록자 교과성적 학생부등급"(원 석차등급 평균, 등급→점수 환산 이전) 그대로이며,
  // 학생 반영교과 원점수 평균과 같은 척도라 시뮬레이터 비교에 바로 쓸 수 있다(가천대와 동일한 방식).
  // 대학환산점수(1,000점 만점)는 참고용으로만 표시하고 계산에는 쓰지 않는다(척도가 달라 직접 비교 불가 — 가천대 사례와 동일 이유).
  const JIGYEOK_2025={"기계공학과":{"canon":"기계공학과","rawName":"기계공학과","recruit":32,"competition":7.1,"prelim":43,"y25cut50":997.37,"y25cut70":997.25,"y26cut50":973.66,"y26cut70":972.53,"rawGrade50":2.29,"rawGrade70":2.36,"track":"N"},"항공우주공학과":{"canon":"항공우주공학과","rawName":"항공우주공학과","recruit":12,"competition":9.7,"prelim":19,"y25cut50":997.45,"y25cut70":997.4,"y26cut50":974.53,"y26cut70":973.98,"rawGrade50":2.23,"rawGrade70":2.3,"track":"N"},"조선해양공학과":{"canon":"조선해양공학과","rawName":"조선해양공학과","recruit":12,"competition":11.6,"prelim":11,"y25cut50":996.71,"y25cut70":996.67,"y26cut50":967.08,"y26cut70":966.67,"rawGrade50":2.62,"rawGrade70":2.63,"track":"N"},"산업경영공학과":{"canon":"산업경영공학과","rawName":"산업경영공학과","recruit":10,"competition":7.8,"prelim":11,"y25cut50":997.13,"y25cut70":997.05,"y26cut50":971.34,"y26cut70":970.53,"rawGrade50":2.43,"rawGrade70":2.46,"track":"N"},"화공에너지공학부(화학공학전공)":{"canon":"화공에너지공학부(화학공학전공)","rawName":"화학공학과","recruit":21,"competition":7.6,"prelim":29,"y25cut50":998.21,"y25cut70":998,"y26cut50":982.12,"y26cut70":980,"rawGrade50":1.89,"rawGrade70":2,"track":"N"},"고분자공학과":{"canon":"고분자공학과","rawName":"고분자공학과","recruit":9,"competition":7.6,"prelim":13,"y25cut50":997.37,"y25cut70":997.2,"y26cut50":973.68,"y26cut70":971.95,"rawGrade50":2.32,"rawGrade70":2.4,"track":"N"},"신소재공학과":{"canon":"신소재공학과","rawName":"신소재공학과","recruit":19,"competition":7.6,"prelim":23,"y25cut50":997.82,"y25cut70":997.59,"y26cut50":978.16,"y26cut70":975.87,"rawGrade50":2.09,"rawGrade70":2.21,"track":"N"},"사회인프라공학과":{"canon":"사회인프라공학과","rawName":"사회인프라공학과","recruit":13,"competition":10.8,"prelim":21,"y25cut50":996.97,"y25cut70":996.85,"y26cut50":969.68,"y26cut70":968.5,"rawGrade50":2.48,"rawGrade70":2.58,"track":"N"},"환경공학과":{"canon":"환경공학과","rawName":"환경공학과","recruit":9,"competition":7.3,"prelim":12,"y25cut50":997.49,"y25cut70":997.34,"y26cut50":974.9,"y26cut70":973.4,"rawGrade50":2.19,"rawGrade70":2.31,"track":"N"},"공간정보공학과":{"canon":"공간정보공학과","rawName":"공간정보공학과","recruit":7,"competition":9.4,"prelim":3,"y25cut50":996.94,"y25cut70":996.93,"y26cut50":969.43,"y26cut70":969.34,"rawGrade50":2.49,"rawGrade70":2.52,"track":"N"},"건축학부":{"canon":"건축학부","rawName":"건축학부","recruit":16,"competition":10.1,"prelim":11,"y25cut50":997.05,"y25cut70":997.01,"y26cut50":970.47,"y26cut70":970.11,"rawGrade50":2.48,"rawGrade70":2.49,"track":"N"},"에너지자원공학과":{"canon":"에너지자원공학과","rawName":"에너지자원공학과","recruit":5,"competition":8.6,"prelim":4,"y25cut50":997.25,"y25cut70":997.13,"y26cut50":972.5,"y26cut70":971.25,"rawGrade50":2.28,"rawGrade70":2.38,"track":"N"},"전기전자공학부":{"canon":"전기전자공학부","rawName":"전기전자공학부","recruit":53,"competition":9.5,"prelim":76,"y25cut50":997.63,"y25cut70":997.33,"y26cut50":976.3,"y26cut70":973.26,"rawGrade50":2.19,"rawGrade70":2.32,"track":"N"},"반도체시스템공학과":{"canon":"반도체시스템공학과","rawName":"반도체시스템공학과","recruit":11,"competition":14.7,"prelim":8,"y25cut50":997.65,"y25cut70":997.49,"y26cut50":976.52,"y26cut70":974.89,"rawGrade50":2.16,"rawGrade70":2.26,"track":"N"},"화공에너지공학부(이차전지공학전공)":{"canon":"화공에너지공학부(이차전지공학전공)","rawName":"이차전지융합학과","recruit":5,"competition":12,"prelim":4,"y25cut50":997.05,"y25cut70":996.63,"y26cut50":970.45,"y26cut70":966.26,"rawGrade50":2.48,"rawGrade70":2.69,"track":"N"},"수학과":{"canon":"수학과","rawName":"수학과","recruit":7,"competition":8.6,"prelim":15,"y25cut50":997.14,"y25cut70":996.91,"y26cut50":971.37,"y26cut70":969.07,"rawGrade50":2.43,"rawGrade70":2.48,"track":"N"},"통계학과":{"canon":"통계학과","rawName":"통계학과","recruit":6,"competition":18,"prelim":6,"y25cut50":996.94,"y25cut70":996.76,"y26cut50":969.44,"y26cut70":967.6,"rawGrade50":2.53,"rawGrade70":2.6,"track":"N"},"물리학과":{"canon":"물리학과","rawName":"물리학과","recruit":8,"competition":13.5,"prelim":6,"y25cut50":997.23,"y25cut70":997.1,"y26cut50":972.29,"y26cut70":971.03,"rawGrade50":2.39,"rawGrade70":2.45,"track":"N"},"화학과":{"canon":"화학과","rawName":"화학과","recruit":10,"competition":9.2,"prelim":13,"y25cut50":997.57,"y25cut70":997.35,"y26cut50":975.71,"y26cut70":973.54,"rawGrade50":2.21,"rawGrade70":2.28,"track":"N"},"해양과학과":{"canon":"해양과학과","rawName":"해양과학과","recruit":7,"competition":6.9,"prelim":5,"y25cut50":997.43,"y25cut70":997.33,"y26cut50":974.29,"y26cut70":973.26,"rawGrade50":2.24,"rawGrade70":2.34,"track":"N"},"식품영양학과":{"canon":"식품영양학과","rawName":"식품영양학과","recruit":9,"competition":7.2,"prelim":7,"y25cut50":996.78,"y25cut70":996.52,"y26cut50":967.79,"y26cut70":965.21,"rawGrade50":2.59,"rawGrade70":2.67,"track":"N"},"경영학과":{"canon":"경영학과","rawName":"경영학과","recruit":30,"competition":7.3,"prelim":74,"y25cut50":997.39,"y25cut70":997.33,"y26cut50":973.88,"y26cut70":973.27,"rawGrade50":2.29,"rawGrade70":2.34,"track":"H"},"파이낸스경영학과":{"canon":"파이낸스경영학과","rawName":"파이낸스경영학과","recruit":7,"competition":9.3,"prelim":11,"y25cut50":996.8,"y25cut70":996.75,"y26cut50":968.04,"y26cut70":967.52,"rawGrade50":2.56,"rawGrade70":2.59,"track":"H"},"아태물류학부":{"canon":"아태물류학부","rawName":"아태물류학부","recruit":15,"competition":6.5,"prelim":31,"y25cut50":997.75,"y25cut70":997.61,"y26cut50":977.52,"y26cut70":976.09,"rawGrade50":2.13,"rawGrade70":2.19,"track":"H"},"국제통상학과":{"canon":"국제통상학과","rawName":"국제통상학과","recruit":14,"competition":8.8,"prelim":18,"y25cut50":997.16,"y25cut70":997.02,"y26cut50":971.58,"y26cut70":970.22,"rawGrade50":2.4,"rawGrade70":2.49,"track":"H"},"국어교육과":{"canon":"국어교육과","rawName":"국어교육과","recruit":5,"competition":8.8,"prelim":21,"y25cut50":998.24,"y25cut70":997.09,"y26cut50":982.38,"y26cut70":970.91,"rawGrade50":1.88,"rawGrade70":2.45,"track":"H"},"영어교육과":{"canon":"영어교육과","rawName":"영어교육과","recruit":5,"competition":18,"prelim":15,"y25cut50":997.95,"y25cut70":997.95,"y26cut50":979.52,"y26cut70":979.52,"rawGrade50":1.98,"rawGrade70":1.98,"track":"H"},"사회교육과":{"canon":"사회교육과","rawName":"사회교육과","recruit":5,"competition":7.2,"prelim":11,"y25cut50":997.67,"y25cut70":997.58,"y26cut50":976.67,"y26cut70":975.79,"rawGrade50":2.17,"rawGrade70":2.21,"track":"H"},"교육학과":{"canon":"교육학과","rawName":"교육학과","recruit":5,"competition":6.2,"prelim":9,"y25cut50":997.89,"y25cut70":996.95,"y26cut50":978.85,"y26cut70":969.48,"rawGrade50":2.06,"rawGrade70":2.48,"track":"H"},"수학교육과":{"canon":"수학교육과","rawName":"수학교육과","recruit":5,"competition":6.2,"prelim":5,"y25cut50":997.73,"y25cut70":997.64,"y26cut50":977.33,"y26cut70":976.43,"rawGrade50":2.13,"rawGrade70":2.18,"track":"N"},"행정학과":{"canon":"행정학과","rawName":"행정학과","recruit":13,"competition":9.8,"prelim":23,"y25cut50":996.98,"y25cut70":996.9,"y26cut50":969.75,"y26cut70":969.03,"rawGrade50":2.47,"rawGrade70":2.54,"track":"H"},"정치외교학과":{"canon":"정치외교학과","rawName":"정치외교학과","recruit":10,"competition":8,"prelim":21,"y25cut50":997.21,"y25cut70":997.02,"y26cut50":972.12,"y26cut70":970.2,"rawGrade50":2.39,"rawGrade70":2.45,"track":"H"},"미디어커뮤니케이션학과":{"canon":"미디어커뮤니케이션학과","rawName":"미디어커뮤니케이션학과","recruit":10,"competition":6.6,"prelim":22,"y25cut50":997.41,"y25cut70":996.96,"y26cut50":974.12,"y26cut70":969.59,"rawGrade50":2.25,"rawGrade70":2.48,"track":"H"},"경제학과":{"canon":"경제학과","rawName":"경제학과","recruit":13,"competition":7.6,"prelim":26,"y25cut50":997.13,"y25cut70":996.94,"y26cut50":971.29,"y26cut70":969.39,"rawGrade50":2.44,"rawGrade70":2.53,"track":"H"},"소비자학과":{"canon":"소비자학과","rawName":"소비자학과","recruit":5,"competition":13,"prelim":7,"y25cut50":996.35,"y25cut70":996.33,"y26cut50":963.46,"y26cut70":963.27,"rawGrade50":2.71,"rawGrade70":2.79,"track":"H"},"아동심리학과":{"canon":"아동심리학과","rawName":"아동심리학과","recruit":5,"competition":6.6,"prelim":12,"y25cut50":997.08,"y25cut70":997.06,"y26cut50":970.77,"y26cut70":970.63,"rawGrade50":2.46,"rawGrade70":2.47,"track":"H"},"사회복지학과":{"canon":"사회복지학과","rawName":"사회복지학과","recruit":5,"competition":7.8,"prelim":9,"y25cut50":996.41,"y25cut70":996.41,"y26cut50":964.1,"y26cut70":964.1,"rawGrade50":2.77,"rawGrade70":2.77,"track":"H"},"한국어문학과":{"canon":"한국어문학과","rawName":"한국어문학과","recruit":8,"competition":8.3,"prelim":10,"y25cut50":996.63,"y25cut70":996.5,"y26cut50":966.32,"y26cut70":965,"rawGrade50":2.66,"rawGrade70":2.7,"track":"H"},"사학과":{"canon":"사학과","rawName":"사학과","recruit":6,"competition":5.3,"prelim":12,"y25cut50":996.96,"y25cut70":996.76,"y26cut50":969.6,"y26cut70":967.6,"rawGrade50":2.49,"rawGrade70":2.55,"track":"H"},"철학과":{"canon":"철학과","rawName":"철학과","recruit":6,"competition":9.3,"prelim":7,"y25cut50":996.76,"y25cut70":996.68,"y26cut50":967.63,"y26cut70":966.75,"rawGrade50":2.59,"rawGrade70":2.61,"track":"H"},"중국학과":{"canon":"중국학과","rawName":"중국학과","recruit":10,"competition":10.5,"prelim":6,"y25cut50":996.83,"y25cut70":996.68,"y26cut50":968.25,"y26cut70":966.81,"rawGrade50":2.59,"rawGrade70":2.66,"track":"H"},"일본언어문화학과":{"canon":"일본언어문화학과","rawName":"일본언어문화학과","recruit":9,"competition":8.4,"prelim":5,"y25cut50":996.76,"y25cut70":996.71,"y26cut50":967.55,"y26cut70":967.07,"rawGrade50":2.57,"rawGrade70":2.62,"track":"H"},"영미유럽인문융합학부":{"canon":"영미유럽인문융합학부","rawName":"영미유럽인문융합학부","recruit":15,"competition":6.9,"prelim":22,"y25cut50":996.95,"y25cut70":996.76,"y26cut50":969.54,"y26cut70":967.6,"rawGrade50":2.52,"rawGrade70":2.62,"track":"H"},"문화콘텐츠문화경영학과":{"canon":"문화콘텐츠문화경영학과","rawName":"문화콘텐츠문화경영학과","recruit":13,"competition":7.2,"prelim":34,"y25cut50":997.17,"y25cut70":996.98,"y26cut50":971.72,"y26cut70":969.75,"rawGrade50":2.4,"rawGrade70":2.51,"track":"H"},"의예과":{"canon":"의예과","rawName":"의예과","recruit":26,"competition":13.6,"prelim":24,"y25cut50":999.81,"y25cut70":999.73,"y26cut50":998.06,"y26cut70":997.33,"rawGrade50":1.1,"rawGrade70":1.13,"track":"N"},"간호학과":{"canon":"간호학과","rawName":"간호학과","recruit":17,"competition":14.4,"prelim":30,"y25cut50":998.06,"y25cut70":998,"y26cut50":980.62,"y26cut70":980,"rawGrade50":1.97,"rawGrade70":2,"track":"N"},"의류디자인학과(일반)":{"canon":"의류디자인학과(일반)","rawName":"의류디자인학과(일반)","recruit":6,"competition":8,"prelim":13,"y25cut50":996.99,"y25cut70":996.84,"y26cut50":969.88,"y26cut70":968.43,"rawGrade50":2.51,"rawGrade70":2.58,"track":"H"},"인공지능공학과":{"canon":"인공지능공학과","rawName":"인공지능공학과","recruit":11,"competition":7.4,"prelim":8,"y25cut50":997.59,"y25cut70":997.41,"y26cut50":975.88,"y26cut70":974.13,"rawGrade50":2.18,"rawGrade70":2.29,"track":"N"},"데이터사이언스학과":{"canon":"데이터사이언스학과","rawName":"데이터사이언스학과","recruit":10,"competition":9.7,"prelim":7,"y25cut50":997.37,"y25cut70":997.34,"y26cut50":973.72,"y26cut70":973.4,"rawGrade50":2.31,"rawGrade70":2.33,"track":"N"},"스마트모빌리티공학과":{"canon":"스마트모빌리티공학과","rawName":"스마트모빌리티공학과","recruit":7,"competition":8.7,"prelim":11,"y25cut50":997.11,"y25cut70":997.06,"y26cut50":971.11,"y26cut70":970.63,"rawGrade50":2.43,"rawGrade70":2.44,"track":"N"},"컴퓨터공학과":{"canon":"컴퓨터공학과","rawName":"컴퓨터공학과","recruit":38,"competition":6.2,"prelim":45,"y25cut50":997.42,"y25cut70":997.05,"y26cut50":974.23,"y26cut70":970.53,"rawGrade50":2.29,"rawGrade70":2.44,"track":"N"},"자유전공융합학부":{"canon":"자유전공융합학부","rawName":"자유전공융합학부","recruit":20,"competition":11.8,"prelim":30,"y25cut50":997.69,"y25cut70":997.58,"y26cut50":976.86,"y26cut70":975.8,"rawGrade50":2.16,"rawGrade70":2.21,"track":"FALL"},"생명공학과":{"canon":"생명공학과","rawName":"생명공학과","recruit":8,"competition":11,"prelim":8,"y25cut50":998.27,"y25cut70":998.16,"y26cut50":982.67,"y26cut70":981.59,"rawGrade50":1.87,"rawGrade70":1.92,"track":"N"},"생명과학과":{"canon":"생명과학과","rawName":"생명과학과","recruit":5,"competition":7.4,"prelim":8,"y25cut50":997.82,"y25cut70":997.82,"y26cut50":978.22,"y26cut70":978.22,"rawGrade50":2.09,"rawGrade70":2.09,"track":"N"}};
  const JONGHAP_2025={"기계공학과":{"canon":"기계공학과","rawName":"기계공학과","recruit":40,"competition":11.7,"prelim":34,"rawGrade50":2.73,"rawGrade70":2.86,"track":"N"},"항공우주공학과":{"canon":"항공우주공학과","rawName":"항공우주공학과","recruit":19,"competition":13.5,"prelim":14,"rawGrade50":2.59,"rawGrade70":2.74,"track":"N"},"조선해양공학과":{"canon":"조선해양공학과","rawName":"조선해양공학과","recruit":19,"competition":10.4,"prelim":4,"rawGrade50":3.43,"rawGrade70":3.57,"track":"N"},"산업경영공학과":{"canon":"산업경영공학과","rawName":"산업경영공학과","recruit":15,"competition":10.4,"prelim":6,"rawGrade50":3,"rawGrade70":3.29,"track":"N"},"화공에너지공학부(화학공학전공)":{"canon":"화공에너지공학부(화학공학전공)","rawName":"화학공학과","recruit":33,"competition":8.6,"prelim":24,"rawGrade50":2.29,"rawGrade70":2.44,"track":"N"},"고분자공학과":{"canon":"고분자공학과","rawName":"고분자공학과","recruit":15,"competition":10.8,"prelim":2,"rawGrade50":2.66,"rawGrade70":2.73,"track":"N"},"신소재공학과":{"canon":"신소재공학과","rawName":"신소재공학과","recruit":32,"competition":10.1,"prelim":27,"rawGrade50":2.57,"rawGrade70":2.7,"track":"N"},"사회인프라공학과":{"canon":"사회인프라공학과","rawName":"사회인프라공학과","recruit":23,"competition":10.8,"prelim":20,"rawGrade50":3.38,"rawGrade70":3.63,"track":"N"},"환경공학과":{"canon":"환경공학과","rawName":"환경공학과","recruit":14,"competition":15.4,"prelim":4,"rawGrade50":2.62,"rawGrade70":2.71,"track":"N"},"공간정보공학과":{"canon":"공간정보공학과","rawName":"공간정보공학과","recruit":12,"competition":10.2,"prelim":4,"rawGrade50":3.44,"rawGrade70":3.73,"track":"N"},"건축학부":{"canon":"건축학부","rawName":"건축학부","recruit":24,"competition":13.5,"prelim":22,"rawGrade50":2.88,"rawGrade70":2.94,"track":"N"},"에너지자원공학과":{"canon":"에너지자원공학과","rawName":"에너지자원공학과","recruit":8,"competition":8.9,"prelim":4,"rawGrade50":2.88,"rawGrade70":2.91,"track":"N"},"전기전자공학부":{"canon":"전기전자공학부","rawName":"전기전자공학부","recruit":71,"competition":9.6,"prelim":57,"rawGrade50":2.55,"rawGrade70":2.68,"track":"N"},"반도체시스템공학과":{"canon":"반도체시스템공학과","rawName":"반도체시스템공학과","recruit":23,"competition":9.3,"prelim":8,"rawGrade50":2.78,"rawGrade70":2.91,"track":"N"},"화공에너지공학부(이차전지공학전공)":{"canon":"화공에너지공학부(이차전지공학전공)","rawName":"이차전지융합학과","recruit":13,"competition":9.5,"prelim":12,"rawGrade50":2.98,"rawGrade70":3.45,"track":"N"},"수학과":{"canon":"수학과","rawName":"수학과","recruit":11,"competition":11.3,"prelim":10,"rawGrade50":2.8,"rawGrade70":2.94,"track":"N"},"통계학과":{"canon":"통계학과","rawName":"통계학과","recruit":9,"competition":10.7,"prelim":5,"rawGrade50":3.05,"rawGrade70":3.2,"track":"N"},"물리학과":{"canon":"물리학과","rawName":"물리학과","recruit":13,"competition":10.2,"prelim":17,"rawGrade50":2.99,"rawGrade70":3.17,"track":"N"},"화학과":{"canon":"화학과","rawName":"화학과","recruit":16,"competition":12.8,"prelim":9,"rawGrade50":2.69,"rawGrade70":2.73,"track":"N"},"해양과학과":{"canon":"해양과학과","rawName":"해양과학과","recruit":11,"competition":15.1,"prelim":2,"rawGrade50":2.73,"rawGrade70":2.83,"track":"N"},"식품영양학과":{"canon":"식품영양학과","rawName":"식품영양학과","recruit":14,"competition":25.6,"prelim":7,"rawGrade50":2.86,"rawGrade70":2.89,"track":"N"},"경영학과":{"canon":"경영학과","rawName":"경영학과","recruit":40,"competition":11.9,"prelim":16,"rawGrade50":2.95,"rawGrade70":3.16,"track":"H"},"파이낸스경영학과":{"canon":"파이낸스경영학과","rawName":"파이낸스경영학과","recruit":11,"competition":9.8,"prelim":5,"rawGrade50":3.18,"rawGrade70":3.23,"track":"H"},"아태물류학부":{"canon":"아태물류학부","rawName":"아태물류학부","recruit":22,"competition":7.4,"prelim":6,"rawGrade50":3.2,"rawGrade70":3.32,"track":"H"},"국제통상학과":{"canon":"국제통상학과","rawName":"국제통상학과","recruit":22,"competition":8.3,"prelim":13,"rawGrade50":3.14,"rawGrade70":3.34,"track":"H"},"국어교육과":{"canon":"국어교육과","rawName":"국어교육과","recruit":8,"competition":7.8,"prelim":6,"rawGrade50":2.49,"rawGrade70":2.54,"track":"H"},"영어교육과":{"canon":"영어교육과","rawName":"영어교육과","recruit":7,"competition":10.1,"prelim":3,"rawGrade50":2.64,"rawGrade70":2.68,"track":"H"},"사회교육과":{"canon":"사회교육과","rawName":"사회교육과","recruit":7,"competition":15.7,"prelim":3,"rawGrade50":2.48,"rawGrade70":2.51,"track":"H"},"교육학과":{"canon":"교육학과","rawName":"교육학과","recruit":7,"competition":11.1,"prelim":8,"rawGrade50":2.58,"rawGrade70":2.62,"track":"H"},"수학교육과":{"canon":"수학교육과","rawName":"수학교육과","recruit":7,"competition":10.3,"prelim":4,"rawGrade50":2.53,"rawGrade70":2.62,"track":"N"},"행정학과":{"canon":"행정학과","rawName":"행정학과","recruit":20,"competition":10.8,"prelim":16,"rawGrade50":2.93,"rawGrade70":3.11,"track":"H"},"정치외교학과":{"canon":"정치외교학과","rawName":"정치외교학과","recruit":14,"competition":19.4,"prelim":10,"rawGrade50":2.9,"rawGrade70":3.05,"track":"H"},"미디어커뮤니케이션학과":{"canon":"미디어커뮤니케이션학과","rawName":"미디어커뮤니케이션학과","recruit":15,"competition":13.1,"prelim":9,"rawGrade50":2.59,"rawGrade70":2.67,"track":"H"},"경제학과":{"canon":"경제학과","rawName":"경제학과","recruit":20,"competition":6.8,"prelim":18,"rawGrade50":3.03,"rawGrade70":3.28,"track":"H"},"소비자학과":{"canon":"소비자학과","rawName":"소비자학과","recruit":7,"competition":11.6,"prelim":1,"rawGrade50":2.98,"rawGrade70":3.07,"track":"H"},"아동심리학과":{"canon":"아동심리학과","rawName":"아동심리학과","recruit":8,"competition":12.6,"prelim":6,"rawGrade50":2.92,"rawGrade70":2.99,"track":"H"},"사회복지학과":{"canon":"사회복지학과","rawName":"사회복지학과","recruit":8,"competition":10.4,"prelim":7,"rawGrade50":2.86,"rawGrade70":2.87,"track":"H"},"한국어문학과":{"canon":"한국어문학과","rawName":"한국어문학과","recruit":12,"competition":9.2,"prelim":9,"rawGrade50":3.3,"rawGrade70":3.34,"track":"H"},"사학과":{"canon":"사학과","rawName":"사학과","recruit":9,"competition":12.7,"prelim":11,"rawGrade50":2.9,"rawGrade70":3.4,"track":"H"},"철학과":{"canon":"철학과","rawName":"철학과","recruit":8,"competition":10,"prelim":9,"rawGrade50":2.99,"rawGrade70":3,"track":"H"},"중국학과":{"canon":"중국학과","rawName":"중국학과","recruit":15,"competition":16.9,"prelim":16,"rawGrade50":3.84,"rawGrade70":4.01,"track":"H"},"일본언어문화학과":{"canon":"일본언어문화학과","rawName":"일본언어문화학과","recruit":13,"competition":12.3,"prelim":10,"rawGrade50":3.84,"rawGrade70":5.09,"track":"H"},"영미유럽인문융합학부":{"canon":"영미유럽인문융합학부","rawName":"영미유럽인문융합학부","recruit":27,"competition":15.9,"prelim":27,"rawGrade50":4.52,"rawGrade70":4.76,"track":"H"},"문화콘텐츠문화경영학과":{"canon":"문화콘텐츠문화경영학과","rawName":"문화콘텐츠문화경영학과","recruit":20,"competition":12.8,"prelim":10,"rawGrade50":2.78,"rawGrade70":2.87,"track":"H"},"의예과":{"canon":"의예과","rawName":"의예과","recruit":42,"competition":16.5,"prelim":27,"rawGrade50":1.13,"rawGrade70":1.41,"track":"N"},"간호학과":{"canon":"간호학과","rawName":"간호학과","recruit":28,"competition":15.8,"prelim":20,"rawGrade50":2.4,"rawGrade70":2.42,"track":"N"},"의류디자인학과(일반)":{"canon":"의류디자인학과(일반)","rawName":"의류디자인학과(일반)","recruit":8,"competition":23.3,"prelim":5,"rawGrade50":3.34,"rawGrade70":3.35,"track":"H"},"인공지능공학과":{"canon":"인공지능공학과","rawName":"인공지능공학과","recruit":22,"competition":9,"prelim":7,"rawGrade50":2.74,"rawGrade70":2.95,"track":"N"},"데이터사이언스학과":{"canon":"데이터사이언스학과","rawName":"데이터사이언스학과","recruit":15,"competition":11.9,"prelim":13,"rawGrade50":2.93,"rawGrade70":3.31,"track":"N"},"스마트모빌리티공학과":{"canon":"스마트모빌리티공학과","rawName":"스마트모빌리티공학과","recruit":12,"competition":9.8,"prelim":6,"rawGrade50":3.09,"rawGrade70":3.14,"track":"N"},"컴퓨터공학과":{"canon":"컴퓨터공학과","rawName":"컴퓨터공학과","recruit":40,"competition":10.2,"prelim":31,"rawGrade50":2.46,"rawGrade70":2.55,"track":"N"},"생명공학과":{"canon":"생명공학과","rawName":"생명공학과","recruit":14,"competition":19.6,"prelim":5,"rawGrade50":2.06,"rawGrade70":2.25,"track":"N"},"생명과학과":{"canon":"생명과학과","rawName":"생명과학과","recruit":11,"competition":26.5,"prelim":8,"rawGrade50":2.34,"rawGrade70":2.45,"track":"N"},"첨단바이오의약학과":{"canon":"첨단바이오의약학과","rawName":"첨단바이오의약학과","recruit":11,"competition":18.9,"prelim":8,"rawGrade50":2.32,"rawGrade70":2.54,"track":"N"}};

  CALCULATION_RULES['inha-schoolrecord-jiyeok-2026']={
    id:'inha-schoolrecord-jiyeok-2026',university:'inha',admissionType:'schoolRecord',appliesToYears:[2026],status:'draft',
    reflectedSubjects:['국어','수학','영어','사회/과학(계열별)'],
    method:'학생부교과(지역균형): 반영교과 인문(국/수/영/사·한국사포함)·자연(국/수/영/과)·자유전공융합학부(전체), 학년-학기 가중치 없음(전 학년 통합 이수단위 가중평균), 진로선택과목은 상위 3과목만 성취도→등급 변환(A=1,B=2,C=4)하여 포함, 석차등급→점수표(2026학년도: 1=10.0~9=0, 진로선택 A/B/C=10/9.8/9.4).',
    notes:['반영교과·진로선택 처리·등급→점수표는 대학어디가 "2026학년도 전형별 주요사항"(Ⅲ-1) 원문으로 확인했다.','반영점수를 1,000점 만점으로 환산하는 마지막 배수 산출식은 이미지로만 제공되어 확인하지 못해 draft로 둔다.','2025학년도 대비 2026학년도에 기본점수 미반영으로 산출식이 변경되어(대학 공지), 두 학년도를 같은 계산규칙으로 비교하지 않는다.'],
    sources:[{org:'대학어디가(adiga.kr)',title:'2026학년도 전형별 주요사항(학생부교과)',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
  };

  const departments=INHA_2025_RESULT_ROWS.map((row,index)=>{
    const [department,group,recruit,competition,extra,cut50,cut70,track,profile70,sourceDepartment]=row;
    const deptId='inha-'+String(index+1).padStart(2,'0');
    const detailId=deptId+'-regular-general';
    const jigyeok=JIGYEOK_2025[department],jonghap=JONGHAP_2025[department];
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
  function jigyeokDetail(deptId,r){
    if(!r)return{details:[]};
    return{details:[{
      id:deptId+'-schoolrecord-jiyeok',name:'학생부교과(지역균형)',group:null,
      years:[{
        year:2025,calculationRuleId:'inha-schoolrecord-jiyeok-2026',
        recruitCount:r.recruit,competitionRatio:r.competition,additionalPass:{count:r.prelim,round:null},
        admissionMethod:'학생부교과 100%(진로선택 상위 3과목 포함, 학년 가중치 없음)',csatMinimum:null,
        indicators:{avgGradeReference50:r.rawGrade50,avgGradeReference70:r.rawGrade70,
          extra:[{label:'대학환산점수(1,000점 만점, 참고용·2025학년도 산식)',value:'50%컷 '+r.y25cut50+' · 70%컷 '+r.y25cut70},
                 {label:'대학환산점수(1,000점 만점, 참고용·2026학년도 산식 재계산)',value:'50%컷 '+r.y26cut50+' · 70%컷 '+r.y26cut70}]},
        sources:[{org:'대학어디가',title:'2025학년도 학생부교과(지역균형) 전형 결과',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
      }]
    }]};
  }
  function jonghapDetail(deptId,r){
    if(!r)return{details:[]};
    return{details:[{
      id:deptId+'-comprehensive-mirae',name:'학생부종합(인하미래인재)',group:null,
      years:[{
        year:2025,calculationRuleId:null,
        recruitCount:r.recruit,competitionRatio:r.competition,additionalPass:{count:r.prelim,round:null},
        admissionMethod:'면접형: 1단계 서류 100%(3.5배수, 의예과 3배수) → 2단계 1단계 70%+면접 30% / 서류형: 서류 100% 일괄(정성평가, 정량화 없음)',csatMinimum:null,
        indicators:{avgGradeReference50:r.rawGrade50,avgGradeReference70:r.rawGrade70,
          extra:[{label:'평가요소',value:'기초학업역량·진로탐구역량·공동체역량(면접형 30/50/20%, 서류형 50/30/20%)'}]},
        sources:[{org:'대학어디가',title:'2025학년도 학생부종합(인하미래인재) 전형 결과',url:INHA_SOURCE_URL,publishedDate:null,checkedDate:'2026-09-02',page:null}]
      }]
    }]};
  }

  // 정시 데이터에 없던 수시 전용 신설/광역 모집단위 2개 추가(영미유럽인문융합학부, 자유전공융합학부[계열 미분리])
  const EXTRA_DEPTS=[
    {canon:'영미유럽인문융합학부',id:'inha-61'},
    {canon:'자유전공융합학부',id:'inha-62'}
  ];
  EXTRA_DEPTS.forEach(({canon,id})=>{
    const jigyeok=JIGYEOK_2025[canon],jonghap=JONGHAP_2025[canon];
    if(!jigyeok&&!jonghap)return;
    departments.push({
      id,name:canon,college:null,
      simulationBinding:{regular:null,schoolRecord:jigyeok?id+'-schoolrecord-jiyeok':null,comprehensive:jonghap?id+'-comprehensive-mirae':null,essay:null},
      admissionTypes:{
        regular:{details:[]},
        schoolRecord:jigyeokDetail(id,jigyeok),
        comprehensive:jonghapDetail(id,jonghap),
        essay:{details:[]}
      }
    });
  });

  UNIVERSITY_DB.push({id:'inha',name:'인하대학교',aliases:['인하대'],region:'인천',guideUrl:INHA_2027_GUIDE_URL,departments});
})();
