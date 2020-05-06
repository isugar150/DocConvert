<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page session="false" %>
<html>
<head>
	<title>Home</title>
</head>
<body>
<form target="_blank" method="post" action="/fileUpload" enctype="multipart/form-data">
	<h1>문서 변환 스프링 테스트</h1>
	<label>파일:</label>
	<input multiple="multiple" type="file" name="file1" accept=".doc,.docx,.hwp,.txt,.html,.xlsx,.xls,.pptx,.ppt,.pdf" />
	<input type="submit" value="upload" />
</form>
<form target="_blank" method="post" action="/webCapture">
	<h1>웹 캡쳐 테스트</h1>
	<label>URL:</label>
	<input type="text" id="URL" name="URL" />
	<input type="submit" value="submit" />
</form>
</body>
</html>
