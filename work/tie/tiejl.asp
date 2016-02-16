<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Write "<div align='center'><font color=blue><font size=-1>靈劍江湖</font></font>"
Response.Write "<font color=red><font size=-1><b>礦石練鐵</b></font></font></div>"
Response.Write "<div align='center'><font color=blue><font size=-1>采礦數量:</font></font>"
Response.Write "<font color=blue><font size=-1><b>已采:<font color=red>"&Session("tiets")&"</font>塊鐵</b></font></font>"
Response.Write "<font color=blue><font size=-1><b>  現有：<font color=red>"&Session("tiesl")&"</font>塊礦石</b></font></font></div>"
if Session("tiesl")<=0 then
	Response.Write "<script Language=Javascript>alert('提示：您現在的礦石已經練完，請選擇賣出換得銀兩，需要再采礦石再練鐵！');</script>"
	response.end
end if
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("tiesj")>22 or session("tiesj")="" then
	session("tiesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
	Session("tiets")=Session("tiets")+1
	Session("tiesl")=Session("tiesl")-2
else
	Response.Write "<script Language=Javascript>alert('嚴重警告，不允許使用鼠標控制程序或速度齒輪！');</script>"
end if
%>
<html>
<head>
<title>礦石練鐵</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
