<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
pass=request.form("pass")
if pass="" then%>
<html>
<head>
<title>�F�C����</title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#660000 text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<center>
<font color="#FFFFFF"><b><font size="+2">�F�C����</font></b></font><br>
<br>
�޲z�n�����f
<form action=login.asp method=post>
�п�J�W�ź޲z�K�X:<input type=password size=20 name=pass style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
    <input type=submit value='�T�{'>
</form>
</center>
</body>
</html>
<%elseif pass=Application("ljjh_adminkey") then
session("ljjh_adminok")=true
Response.Redirect "admin.asp"
else
Response.Redirect "../exit.asp"
end if%> 