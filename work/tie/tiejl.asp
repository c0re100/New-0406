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
Response.Write "<font color=red><font size=-1><b>�q�۽m�K</b></font></font></div>"
Response.Write "<div align='center'><font color=blue><font size=-1>���q�ƶq:</font></font>"
Response.Write "<font color=blue><font size=-1><b>�w��:<font color=red>"&Session("tiets")&"</font>���K</b></font></font>"
Response.Write "<font color=blue><font size=-1><b>  �{���G<font color=red>"&Session("tiesl")&"</font>���q��</b></font></font></div>"
if Session("tiesl")<=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�z�{�b���q�ۤw�g�m���A�п�ܽ�X���o�Ȩ�A�ݭn�A���q�ۦA�m�K�I');</script>"
	response.end
end if
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("tiesj")>22 or session("tiesj")="" then
	session("tiesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
	Session("tiets")=Session("tiets")+1
	Session("tiesl")=Session("tiesl")-2
else
	Response.Write "<script Language=Javascript>alert('�Y��ĵ�i�A�����\�ϥι��б���{�ǩγt�׾����I');</script>"
end if
%>
<html>
<head>
<title>�q�۽m�K</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#CADEFD" text="#000000" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
