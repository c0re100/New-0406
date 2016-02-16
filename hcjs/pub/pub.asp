<%
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<link rel="stylesheet" href="../../style.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖客棧</title></head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="#000000">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;; width: 237px; height: 16px"></DIV><SCRIPT LANGUAGE="JavaScript" src="../../ljimage/lj.js"></SCRIPT> 
<p align="center" style="font-size:16;color:yellow"><IMG SRC="bg.jpg" WIDTH="400" HEIGHT="300" USEMAP="#Map" BORDER="0"> 
<MAP NAME="Map"><AREA SHAPE="rect" COORDS="259,118,349,201" HREF="pub0.asp"  ONMOUSEOVER="drs('歡迎來到客棧，要休息嗎！'); return true;" ONMOUSEOUT="nd(); return true;"></MAP> 
</body>

</html>
