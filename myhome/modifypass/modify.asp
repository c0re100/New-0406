<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

 
<!--#include file="../data.asp"-->
<%if info(0)<>"" then
name=info(0)
end if%>

<%Response.buffer=true
oldpass=Request.form("oldpass")
pass=request.form("pass")
repass=request.form("repass")
repass=trim(repass)
if instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then
response.end
end if
if instr(oldpass,"'")<>0 or instr(oldpass,"|")<>0 or instr(oldpass," ")<>0 then
response.end
end if
if instr(pass,"'")<>0 or instr(pass,"|")<>0 or instr(pass," ")<>0 then
response.end
end if
if instr(repass,"'")<>0 or instr(repass,"|")<>0 or instr(repass," ")<>0 then
response.end
end if


if oldpass="" or pass="" or repass="" then
message="�K�X���ର��"

elseif pass<>repass then
message="�K�X�P�T�{�K�X�����@�P"
else
'ip=Request.ServerVariables("REMOTE_ADDR")
%>
<%

'����Τ�
sql="SELECT * FROM �Τ� WHERE �m�W='" & name & "' and �K�X='" & oldpass& "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="�藍�_�A�A���K�X����"
else
if  rs("���A")="�v" then
message="�A�{�b�𮧤��A����ק�K�X�I"
else
sql="update �Τ� set �K�X='" & pass & "' where �m�W='" & name & "'"
conn.Execute(sql)
message="���߱z���\�a�ק�F�K�X�I"
end if
end if
conn.close
end if
%>

<head>
<style>td{font-size:14}</style>
<title></title>
</head>

<body>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>


</td></tr>
</table>
</td></tr>
</table>
</body>