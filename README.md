# DocConvert
* WebSocket으로 통신하여 오피스파일(docx, doc, xlsx, xls, pptx, ppt), 한글파일(hwp), 기타 파일(pdf, txt, html)등 여러 파일을 PDF로 변환하거나 이미지 파일(jpg, png, bmp)로 변환합니다.
* 기본포트: webSocket: 12000, FTPServer: FTP/FTPS 12100/12150
* 이미지 변환 ENUM: (0: 변환안함), (1: JPG), (2: PNG), (3: BMP)  
* 권장 오피스 버전: 2010 이상
* 권장 한글 버전: 2010 이상 - 마이너 업데이트 필수 (https://www.hancom.com/cs_center/csDownload.do)
* 권장 운영체제: Windows 10이상, Windows Server 2016이상
* 작업 공간폴더, 로그 자동정리 스케줄러 추가 (Default: 매일 00시 실행)  
* 웹 캡쳐 두가지 방식  
  * PhantomJS  
    * 빠르고 가벼우나 캡쳐가 완벽하게 안됨.
  * CefSharp (Chromium 기반)
    * 크로미움 브라우저로 거의 완벽하게 캡쳐.

## 설치 가이드
1. 압축 해제후 DocConvert_Server.ini 환경에 맞게 설정
2. DocConvert_Util 실행 후 HWP DLL 등록 체크
3. Install 폴더안에 있는 FileZilla Server 인스톨
4. FileZilla Server Interface에서 Edit - Setting에서 Listen on these ports: 12100 입력
5. Edit - Users에서 Add-유저이름입력 Password지정 왼쪽 탭에 Shared folders클릭 후 Add - 변환 폴더 지정, 오른쪽 Files와 Directories AllCheck후 확인
6. Convert_Util로 변환되는지 확인.

[![Video Label](http://www.namejm.org/resources/img/20200408_DEMO_000.jpg)](https://youtu.be/AfrIzDilIZo)  
**<데모 동영상>**

![dcLogic](https://user-images.githubusercontent.com/13088077/78665631-16137400-7911-11ea-8843-5320c42fa519.png)   

> 사용된 C#라이브러리: vtortola.WebSockets, websocket-sharp, pdfium, NLOG, OfficeAPI, HWP API, log4j, FluentFTP, phantomjs    
> 사용된 Java라이브러리: commons-net-3.6.jar, Java-WebSocket-1.4.1.jar, json-simple-1.1.1.jar, slf4j-api-1.7.25.jar  
## 프로그램별 설명
* 설정한 포트로 WebSocket을 Listen하며 요청이 들어오면 해당 문서를 변환해주는 역할  
![image](https://user-images.githubusercontent.com/13088077/80710856-635abe00-8b2a-11ea-9cf6-db09143aac8b.png)  
<DocConvert 서버 메인화면>  
* 혹시라도 서버 프로그램이 오류로인해 강제로 종료될경우 자동으로 실행시켜주는 관리 프로그램  
![image](https://user-images.githubusercontent.com/13088077/80711349-2e02a000-8b2b-11ea-8dd2-0bc5e59731f0.png)  
<DocConvert 매니저 메인화면>  
* 서버가 정상 동작하는지 가장빠르게 확인할수있는 프로그램  
![image](https://user-images.githubusercontent.com/13088077/80710846-5f2ea080-8b2a-11ea-86f0-030d05112780.png)  
<DocConvert 유틸 메인화면>  
* 자바 개발시 서버에 요청하고 다운로드받는 작업을 도와주는 API
![image](https://user-images.githubusercontent.com/13088077/80711207-f136a900-8b2a-11ea-869e-1d4eb244a901.png)  
<JAVA API를 사용하여 변환시 CentOS 환경>  

### DocConverter Server
**Socket & WebSocket**  
`문서 변환시`
```
- [Client => Server]  
  "KEY": "ANY",  
  "FileName": "DOCUMENT.xlsx",  
  "ConvertIMG": 2,  
  "DocPassword": "",  
  "useCompression": true //압축 다운로드시 사용  
    
 - [Server => Client]  
  "convertImgCnt": "92",  
  "URL": "/workspace/74C5728223BB8AE604AC1056C0D7DC2A",  
  "isSuccess": true,  
  "msg": "변환에 성공하였습니다.",  
  "zipURL": "/workspace/74C5728223BB8AE604AC1056C0D7DC2A/DOCUMENT.zip" // 압축파일 경로  
```  
`웹 캡쳐시`
```
- [Client => Server]  
  [Socket][Client => Server]
  "KEY": "Any",
  "Method": "WebCapture",
  "URL": "http://www.naver.com",
    
 - [Server => Client]  
  "URL": "/workspace/0d73a975-b29b-463f-85a1-44d61c91508c_2020-04-16/ImageName.png",
  "isSuccess": true,
  "msg": "WebCapture에 성공하였습니다.",
  "convertImgCnt": 1,
  "Method": "WebCapture"
```  

## DocConverter Engien
```
WordConvert_Core.WordSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  
```
```
ExcelConverter_Core.ExcelSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  
```
```
PowerPointConvert_Core.PowerPointSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  
```
```
HWPConvert_Core.HwpSaveAs(String FilePath, String outPath, bool PageCounting)
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
bool pageCounting: 변환 전 페이지번호 추출 여부  
```
```
ConvertImg.PDFtoJpeg(String SourcePDF, String outPath)
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  
```
```
ConvertImg.PDFtoBmp(string SourcePDF, string outPath)
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  
```
```
ConvertImg.PDFtoPng(string SourcePDF, string outPath)
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  
```
```
WebCapture_Core.WebCapture(string Url, string outPath)
String SourcePDF: 웹 캡쳐 URL  
String outPath: 이미지 내보낼 경로  
```
## FAQ
```
Q: 동시에 몇개까지 변환이 가능합니까?
A: 서버 사양이 따라주면 동시변환 제한은 따로 없습니다.
```
```
Q: 최대 변환 가능한 페이지수는 몇장인가요?
A: 정확하게 측정은 안해봤으나 1000장 이상 변환되는걸 확인했습니다.
```
```
Q: 한글에서 PDF 변환 시 오른쪽과 하단에 공백이 생깁니다.
A: 한글 처음 설치 후 바로 변환하면 해당 증상이 발생하며 해결 방법은 임의 문서를 PDF로 인쇄하고 다시시도하면 정상작동합니다.
```
```
Q: 한글 변환시 프로그램이 꺼집니다.
A: 해당 문제는 한글 버전이 낮아서 발생하며, 구버전 한글의 경우 마이너 업데이트를 진행해야 정상작동합니다.
```
## License
License 파일 참조
