# SundayWorshipPPTMaker
주일예배 PPT 자동화 작업용 툴.

주일예배 ppt는 고정된 몇 개의 슬라이드에 성경 구절, 광고, 교역자로부터 배포되는 찬양가사ppt와 설교ppt, 관련 영상 등을 추가하는 작업을 거쳐 만든다.
작업에 필요한 정보는 함께 배포되는 주보 서식 한글(.hwp)파일을 통해 확인한다.

파일들을 선택하면 아래 루틴에 대해 자동화 작업을 수행하며 작업 편의에 따라 몇가지 순서가 변경된다.
본 툴은 작성자의 케이스에 최적화되어 있다.

	1. 주보 확인(찬양 순서, 대표기도자, 말씀범위, 설교제목, 광고, 생일자)
	2. ppt템플릿 파일(주마다 변경되지 않는 고정된 슬라이드만 있음)을 연다.
	3. 찬양 가사ppt 삽입
	4. 찬양 제목 슬라이드 작성
	5. 영상, 설교 ppt 삽입
	6. 광고(생일자 포함) 작성
	7. 말씀 본문 작성

## Workflow


----------------------------------------------
## Requirements
* Visual Studio 2019 (한글 COM은 32bit VS2022는 64bit이라 사용 불가 )
* .NET 5.0 or later

## Dependency
* Microsoft.Office.Core = 15.0.0
* Microsoft.Office.Interop.PowerPoint = 15.0.4420.1028
* HwpObjectLib = 1.0.0	(HwpCtrl api는 오류 발생)
* [한글 보안 모듈](https://www.hancom.com/board/devdataView.do?board_seq=47&artcl_seq=4084&pageInfo.page=&search_text=)
* WindowsApiCodePack = 1.1.3 (not essential)

## Others
* 성경 축약어 텍스트
* 한글 성경 데이터
--------------------------------------------------
## 
![녹화_2022_09_03_21_29_47_277](https://user-images.githubusercontent.com/32349691/197352259-ed067af2-3991-4e7e-8b96-94fee34a4d03.gif)

## 22.10.31
![image](https://user-images.githubusercontent.com/32349691/199148140-ee914feb-d59d-4640-a1c2-a0e2b08789de.png)
![image](https://user-images.githubusercontent.com/32349691/199148518-a19ce9bb-780f-48e5-ab37-0e3b2dce1e85.png)
![image](https://user-images.githubusercontent.com/32349691/199148693-c6027540-5d7b-43dc-b0bd-28197927acff.png)
1. 설정 창 추가 & config파일 연동, Help Popup
2. 로고, 아이콘 추가
3. 버튼, 레이아웃 변경
4. 랜덤 커버 변경

## 22.12.26
![image](https://user-images.githubusercontent.com/32349691/209490068-922a8c26-144c-49cd-a57e-7c52a2c90818.png)
1. DarkMode Style
2. 주보 파일 열수 없을 때 manual 설정
3. 범위 텍스트 정규식
4. 클래스, 함수 분리

## Next..
1. Calendar Pdf Parsing
2. 가사ppt 스크래핑
3. Web연동
