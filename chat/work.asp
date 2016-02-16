<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>職業</title>
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
<p>此項功能保留，有待開發！<br>
<font color="#000000">職業選擇</font><br>
<a href="#"onClick="window.open('../work/ice/icemain.asp','caibing','scrollbars=yes,resizable=yes,width=700,height=420')" title="采冰賺錢！">采 冰</a><br>
<a href="#"onClick="window.open('../work/mine/minemain.asp','caikuang','scrollbars=yes,resizable=yes,width=700,height=420')" title="采冰賺錢！">采 礦</a>
<a href="#"onClick="window.open('../work/tie/tiemain.asp','liantie','scrollbars=yes,resizable=yes,width=700,height=420')" title="采冰賺錢！">練 鐵</a>
<p>不同的職業可以有不同的收入，當然像道德、內力、體力、魅力等也是不相同的，大家可以根據自己的喜好來選擇自己的職業！<br>
</p>
</div>
</body>
</html>
 