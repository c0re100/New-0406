<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Write "<div align='center'><font color=blue><font size=-1>�F�C����</font></font>"
Response.Write "<font color=red><font size=-1><b>��s���</b></font></font></div>"
Response.Write "<div align='center'><font color=blue><font size=-1>���Y�ƶq:</font></font>"
Response.Write "<font color=blue><font size=-1><b>�w��:<font color=red>"&Session("minets")&"</font>�ڤ��Y</b></font></font>"
Response.Write "<font color=blue><font size=-1><b>  �{���G<font color=red>"&Session("minesl")&"</font>�ڤ��Y</b></font></font></div>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("minesj")>22 or session("minesj")="" then
	session("minesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
	Session("minets")=Session("minets")+1
else
	Response.Write "<script Language=Javascript>alert('�Y��ĵ�i�A�����\�ϥι��б���{�ǩγt�׾����I');</script>"
end if
%>
<html>
<head>
<title>��s���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
