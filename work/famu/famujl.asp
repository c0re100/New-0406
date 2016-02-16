<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Write "<div align='center'><font color=blue><font size=-1>艶糃打</font></font>"
Response.Write "<font color=red><font size=-1><b>ワれ</b></font></font></div>"
Response.Write "<div align='center'><font color=blue><font size=-1>れ繷计秖:</font></font>"
Response.Write "<font color=blue><font size=-1><b>:<font color=red>"&Session("minets")&"</font>れ繷</b></font></font>"
Response.Write "<font color=blue><font size=-1><b>  瞷Τ<font color=red>"&Session("minesl")&"</font>れ繷</b></font></font></div>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("minesj")>22 or session("minesj")="" then
	session("minesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
	Session("minets")=Session("minets")+1
else
	Response.Write "<script Language=Javascript>alert('腨牡ぃす砛ㄏノ公夹北祘┪硉睛近');</script>"
end if
%>
<html>
<head>
<title>ワれ</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
