<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Write "<div align='center'><font color=blue><font size=-1>�F�C����</font></font></div>"
Response.Write "<br>"
Response.Write "<div align='center'><font color=red><font size=-1><b>���a���B</b></font></font></div>"
Response.Write "<br>"
Response.Write "<font color=blue><font size=-1>����:</font></font><img src='images/tong"& Session("icejl1") &".gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<font color=blue><font size=-1>���B:</font></font><img src='images/tong4.gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=blue><font size=-1><b>�w��:"&Session("icets")&"��</b></font></font>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=blue><font size=-1><b>�{���G"&Session("icesl")&"��</b></font></font>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("cbsj")>12 or session("cbsj")="" then
session("cbsj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
Session("icejl1")=Session("icejl1")+1
if Session("icejl1")=4 then
Session("icejl1")=0
Session("icets")=Session("icets")+1
end if
else
Response.Write "<script Language=Javascript>alert('�Y��ĵ�i�A�����\�ϥΪ��B���ܳt�׾����I');</script>"
end if
%>
<html>
<head>
<title>���a���B</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>