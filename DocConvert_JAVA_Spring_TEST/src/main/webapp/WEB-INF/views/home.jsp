<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page session="false" %>
<html>
<head>
	<title>Home</title>
</head>
<body>
<form method="post" action="fileUpload" enctype="multipart/form-data">
	<h1>DocConvert_Spring Test</h1>
	<label>파일:</label>
	<input multiple="multiple" type="file" name="file1">
	<input type="submit" value="upload">
</form>
</body>
</html>
