<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>¾�~</title>
<style type="text/css">
<!--
body {  font-size: 9pt}
a:link {  text-decoration: none}
a:hover {  text-decoration: underline}
-->
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background=bk.jpg bgproperties="fixed" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p>�����\��O�d�A���ݶ}�o�I<br>
<font color="#000000">¾�~���</font><br>
<a href="#"onClick="window.open('../work/ice/icemain.asp','caibing','scrollbars=yes,resizable=yes,width=700,height=420')" title="���B�ȿ��I">�� �B</a><br>
<a href="#"onClick="window.open('../work/mine/minemain.asp','caikuang','scrollbars=yes,resizable=yes,width=700,height=420')" title="���B�ȿ��I">�� �q</a>
<a href="#"onClick="window.open('../work/tie/tiemain.asp','liantie','scrollbars=yes,resizable=yes,width=700,height=420')" title="���B�ȿ��I">�m �K</a>
<p>���P��¾�~�i�H�����P�����J�A��M���D�w�B���O�B��O�B�y�O���]�O���ۦP���A�j�a�i�H�ھڦۤv���ߦn�ӿ�ܦۤv��¾�~�I<br>
</p>
</div>
</body>
</html>
 