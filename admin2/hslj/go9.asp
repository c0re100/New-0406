<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache" 
if Session("iq")<>"ok" then
response.redirect "index.asp"
end if
my=ljjh_nickname
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
sql="SELECT ���O FROM �Τ� where �m�W='" & my & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
Response.Write "<script language=JavaScript>{alert('�A���O�F�C���򤤤H�I');location.href = 'index.asp';}</script>"
set rs=nothing
conn.close
set conn=nothing
Response.End
else
at=rs("���O")
if ljjh_nickname="" then
Response.Write "<script language=JavaScript>{alert('�A���O�F�C���򤤤H�I');location.href = 'index.asp';}</script>"
set rs=nothing
conn.close
set conn=nothing
Response.End
else
Randomize
at=int(at*rnd)
Randomize
at1=int(5000000*rnd)
if at>at1 then
Session("iq")="hla"
msg="�g�L�@�f�W����A�A�H"&at&"�I���OĹ�F�L������"&at1&"�I���O�A<a href='go10.asp'>���o�̥Τ��O�_�}�_�ê����a�C</a>"
else
Session("iq")=""
msg="�g�L�@�f�W����A�A�H"&at&"�I���O���ĪL������"&at1&"�I���O�A<a href='javascript:self.close()'>���o�̦^�h�a�C</a>"
end if
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title>��Z���G</title><LINK href="../style.css" rel=stylesheet></head>
<body background="images/bg.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"><%=msg%></div>
</html>
<%
end if
end if
%>
