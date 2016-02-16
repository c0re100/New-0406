<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Write "<div align='center'><font color=blue><font size=-1>靈劍江湖</font></font></div>"
Response.Write "<br>"
Response.Write "<div align='center'><font color=red><font size=-1><b>極地采冰</b></font></font></div>"
Response.Write "<br>"
Response.Write "<font color=blue><font size=-1>完成:</font></font><img src='images/tong"& Session("icejl1") &".gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<font color=blue><font size=-1>采冰:</font></font><img src='images/tong4.gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=blue><font size=-1><b>已采:"&Session("icets")&"桶</b></font></font>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=blue><font size=-1><b>現有："&Session("icesl")&"桶</b></font></font>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("cbsj")>12 or session("cbsj")="" then
session("cbsj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
Session("icejl1")=Session("icejl1")+1
if Session("icejl1")=4 then
Session("icejl1")=0
Session("icets")=Session("icets")+1
end if
else
Response.Write "<script Language=Javascript>alert('嚴重警告，不允許使用采冰機變速度齒輪！');</script>"
end if
%>
<html>
<head>
<title>極地采冰</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>