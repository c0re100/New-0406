<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("iq")<>"opy" then
response.redirect "index.asp"
end if
%>
<%
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT ���� FROM �Τ� where id="&info(9)
Set Rs=conn.Execute(sql)
at=rs("����")
randomize timer
r=int(rnd*1)
s=int(rnd*5)
at=at*s
randomize timer
r=int(rnd*1)
s=int(rnd*5)
at1=100000*s
if at>at1 then
Session("iq")="ok"
msg="�g�L�@�f�W����A�A�H"&at&"�I����Ĺ�F��������"&at1&"�I�����A<a href='go6.asp'>���o�����U�@���a�C</a>"
else
Session("iq")=""
msg="�g�L�@�f�W����A�A�H"&at&"�I�������ĥ�������"&at1&"�I�����A<a href='javascript:self.close()'>���o�̦^�h�a�a�C</a>"
end if
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title>��Z���G</title>
</head>
<body bgcolor="#0099FF" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<div align="center"><b><%=msg%></b></div></body>
</html>