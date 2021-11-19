## 크롤링 연습

---

##### 문제 >>  
Step1. 국민청원 사이트에서 '코로나'를 검색하여 들어가기  
Step2. Title, Content, Start, End, Amount, Link 값들을 가져오기  
Step3-1. 현재 날짜로부터 10전 청원까지 가져오기  
Step3-2. 청원인원이 1,000명 이상인 데이터만 가져오기  
Step4. 가져온 청원 데이터들을 Excel 안에 넣어주자  
Step5. 저장한 데이터를 Email로 보내주자

---


##### Step1. 국민청원 사이트에서 '코로나'를 검색하여 들어가기  

1. `Open Browser` or `Launch Website` 을 이용하여 인터넷을 실행
   - Open Browser : Internet Explorer 실행.
   - Launch Website : Window 설정에서 기본 앱으로 등록된 웹 브라우저 실행
2. `Object Cloning` 기능의 **Capture**을 사용하여 Window 오브젝트, 엘리먼트 속성들을 컨트롤 할 수 있다.
   - (크롬 개발자 도구를 활용하면 더욱 편리하게 사용이 가능하다.)
   - `Ctrl + Shift + C` 를 사용하면 Element 값인 `xPath` 값 을 쉽게 복사 가능
3. `Keystrokes` 기능을 이용하여 키보드 키 값의 입력을 수행한다.
   - `Mouse Click` 은 좌표 값이며 불완전 하므로  `Object Cloning` 와 `Keystrokes` 기능을 중점으로 사용하도록 하자.

![image](https://user-images.githubusercontent.com/79305451/140871479-1fe39e02-f595-49ab-a8fe-314077860123.png)


##### Step2. Title, Content, Start, End, Amount, Link 값들을 가져오기  

##### Step2-1. 파일 설정
1. 파일 경로를 \$vFileName\$명으로 변수 설정 해준다.
2. `IF문` 을 통해 실행 파일이 존재 하지 않으면 임시파일을 카피하여 새로 생성해준다.
3. Excel 의 Open Spreadsheet 를 열고 Set Cell 을 통해 원하는 위치에 값을 적어 헤더 값을 적어준다.
   - Active Cell 자동 설정 , Specific Cell 특정 셀의 위치 설정
4. 값을 저장하고 종료 해준다.

![image](https://user-images.githubusercontent.com/79305451/140871384-0e4b47aa-3156-451b-8a6b-91f95e650cea.png)

##### Step2-2. 값 가져오기
1. 한 페이지에 총 10개의 청원글이 등록되어 있으므로 Loop는 10Times 동작하도록 한다.
2. `Ctrl + Shift + C` 를 이용하여 각각의 xPath값들을 가져와 `Object Cloning` 한다.
3. xPath 값은 반복적인 데이터 이므로 시스템 변수 \$Counter\$를 이용하여 각각의 데이터를 추출할 수 있도록 설정한다.

##### Step3. 각각의 조건에 부합하는 값들만 가져오기

##### Step3-1. 현재 날짜로부터 10일 전 청원까지만 가져오기  
1. 일자를 비교하기 위하여 \$vDateWrite\$ (작성일자), \$vDateCondition\$ (기한일) 두개의 사용자 변수를 생성해준다.
2. 작성일자는 Step2-2 에서 추출한 \$vColumn3\$ 값을 그대로 넣어준다.
3. 기한일은 시스템 변수인 \$Date\$ 를 이용하여 -10 을 적용하여 포매팅 한다.
##### Step3-2. 청원인원이 1,000명 이상인 데이터만 가져오기  
1. If문을 활용하여 조건을 등록해준다.
2. 청원인원은 Step2-2 에서 추출한 \$vColumn5\$ 값과 바로 비교해 준다.
3. 조건에 부합하면 Continue 를 이용해 저장하지 않고 다음 단계로 넘어가도록 한다.

![image](https://user-images.githubusercontent.com/79305451/140882926-0142786e-7810-4fca-9831-71fd47d9adcd.png)


##### Step4. 가져온 청원 데이터들을 Excel 안에 넣어주자  

1. Step2-1 에서 Excel 사용한 방식과는 다르게 데이터베이스 기능을 이용해 보자
   1. 속도 측면에서 훨씬 빠르고 간결하게 사용이 가능하다
2. 우선 INSERT 해줄 데이터베이스를 연결 해주어야 한다.
3. Database 의 INSERT 를 이용하여 데이터를 넣어준다.
   1. INSERT INTO [Sheet1\$] (\[title],\[content],\[start],\[end],\[amount],\[link])
VALUES ("\$vColumn1\$","\$vColumn2$","\$vColumn3$",
"\$vColumn4$","\$vColumn5$","\$vColumn6$");

![image](https://user-images.githubusercontent.com/79305451/140892522-a91a0d76-f3b3-4522-925c-1fc2b868b126.png)


##### Step5. 저장한 데이터를 Email로 보내주자
