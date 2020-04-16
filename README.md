# DocConvert
* Socket 또는 WebSocket으로 통신하여 오피스파일(docx, doc, xlsx, xls, pptx, ppt), 한글파일(hwp), 기타 파일(pdf, txt, html)등 여러 파일을 PDF로 변환하거나 이미지 파일(jpg, png, bmp)로 변환합니다.
* 기본포트: Socket: 12000, webSocket: 12005, FTPServer: 12100  
* 이미지 변환 ENUM: (0: 변환안함), (1: JPG), (2: PNG), (3: BMP)  
* 권장 오피스 버전: 2016이상
* 권장 운영체제: Windows 10이상, Windows Server 2016이상

### 설치 가이드
1. 압축 해제후 DocConvert_Server.exe.config 환경에 맞게 설정
2. DocConvert_Util 실행 후 HWP DLL 등록 체크
3. Install 폴더안에 있는 FileZilla Server 인스톨
4. FileZilla Server Interface에서 Edit - Setting에서 Listen on these ports: 12100 입력
5. Edit - Users에서 Add-유저이름입력 Password지정 왼쪽 탭에 Shared folders클릭 후 Add - 변환 폴더 지정, 오른쪽 Files와 Directories AllCheck후 확인
6. Convert_Util로 변환되는지 확인.

[![Video Label](http://www.namejm.org/resources/img/20200408_DEMO_000.jpg)](https://youtu.be/AfrIzDilIZo)  
**<데모 동영상>**

![dcLogic](https://user-images.githubusercontent.com/13088077/78665631-16137400-7911-11ea-8843-5320c42fa519.png)   

> 사용된 C#라이브러리: SuperSocket, vtortola.WebSockets, pdfium, NLOG, OfficeAPI, HWP API, log4j, FluentFTP, phantomjs    
> 사용된 Java라이브러리: commons-net-3.6.jar, Java-WebSocket-1.4.1.jar, json-simple-1.1.1.jar, slf4j-api-1.7.25.jar  
### 스크린샷

![dcs_main](https://user-images.githubusercontent.com/13088077/78530590-e7b56c00-781e-11ea-81e2-e0b174e1773e.png)  
<DocConvert 서버 메인화면>    
<img width="487" alt="dcu_main" src="https://user-images.githubusercontent.com/13088077/78663785-b49dd600-790d-11ea-8f60-ef3e9048062b.png">  
<DocConvert 유틸 메인화면>    
<img width="1136" alt="dcjavaapi" src="https://user-images.githubusercontent.com/13088077/78663993-16f6d680-790e-11ea-8afd-0d31d5ec9943.png">  
<JAVA API를 사용하여 변환시>  

### DocConverter Server
**Socket & WebSocket**
`문서 변환시`
```
- [Client => Server]  
  "KEY": "ANY",  
  "FileName": "DOCUMENT.xlsx",  
  "ConvertIMG": 2,  
  "DocPassword": ""  
    
 - [Server => Client]  
  "convertImgCnt": "92",  
  "URL": "/workspace/74C5728223BB8AE604AC1056C0D7DC2A",  
  "isSuccess": true,  
  "msg": "변환에 성공하였습니다."  
```  
`웹 캡쳐시`
```
- [Client => Server]  
  [Socket][Client => Server]
  "KEY": "B29D00A3 - F825 - 4EB7 - 93C1 - A77F5E31A7C2",
  "Method": "WebCapture",
  "URL": "http://www.naver.com"
    
 - [Server => Client]  
  "URL": "/workspace/0d73a975-b29b-463f-85a1-44d61c91508c_2020-04-16/ImageName.png",
  "isSuccess": true,
  "msg": "WebCapture에 성공하였습니다.",
  "convertImgCnt": 1,
  "Method": "WebCapture"
```  

### DocConverter Engien
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
### License
