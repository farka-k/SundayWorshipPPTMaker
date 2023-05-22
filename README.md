# SundayWorshipPPTMaker
주일예배 PPT 자동화 작업용 툴.

주일예배 ppt는 고정된 몇 개의 슬라이드에 성경 구절, 광고, 교역자로부터 배포되는 찬양가사ppt와 설교ppt, 관련 영상 등을 추가하는 작업을 거쳐 만든다.
작업에 필요한 정보는 함께 배포되는 주보 서식 한글(.hwp)파일을 통해 확인한다.

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
* System.Data.SQLite = 1.0.116
* HtmlAgilityPack = 1.11.46
* OpenCVSharp = 4.7.0.X
* Tesseract = 5.2.0

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

## ~2023 May
1. Adapt Clova OCR
2. Code based template
3. Setting Option 수정
4. 광고 영역 적용
5. 가사 검색 관련 UI,기능 추가 

## Next
1. 가사 검색
2. Custom Lyric Slide
3. Tesseract OCR + OpenCV
4. Multithread
5. 역본 DB추가
