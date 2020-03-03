# 소개
# DocConverter Server
**Socket**
- 준비중

# DocConverter Engien
- 지원하는 프로그램(Word, Excel, PowerPoint, HWP)
- PDF TO IMAGE변환 기능 지원


**WordConvert_Core.WordSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)**  
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  

**ExcelConverter_Core.ExcelSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)**  
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  

**PowerPointConvert_Core.PowerPointSaveAs(String FilePath, String outPath, String docPassword, bool pageCounting, bool appvisible)**  
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
String docPassword: 해당 문서의 비밀번호(없으면 비워두면됨)  
bool pageCounting: 변환 전 페이지번호 추출 여부  
bool appvisible: 변환 진행상황을 표시할지 여부  

**HWPConvert_Core.HwpSaveAs(String FilePath, String outPath, bool PageCounting)**  
String FilePath: 소스 파일경로  
String outPath: PDF로 내보낼 파일경로  
bool pageCounting: 변환 전 페이지번호 추출 여부  

**ConvertImg.PDFtoJpeg(String SourcePDF, String outPath)**  
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  

**ConvertImg.PDFtoBmp(string SourcePDF, string outPath)**  
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  

**ConvertImg.PDFtoPng(string SourcePDF, string outPath)**  
String SourcePDF: 타겟 PDF파일  
String outPath: 이미지 내보낼 경로  
