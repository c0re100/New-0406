<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM ���~�R WHERE ���~�W='" & name & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
if clng(Request.Form("wupinll"))<0 then Response.Redirect "../error.asp?id=471"
if clng(Request.Form("wupintl"))<0 then Response.Redirect "../error.asp?id=471"
name=Request.Form("wupinname")
say=Request.Form("wupinsay")
a=clng(Request.Form("wupintl"))
b=clng(Request.Form("wupinfy"))
c=clng(Request.Form("wupingj"))
d=clng(Request.Form("wupinjg"))
e=clng(Request.Form("wupinnl"))
wupinll=(a+b)*400
set rs=server.createobject("adodb.recordset")
sql="select * from ���~�R where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("���~�W")=name
rs("����")=request.form("lx")
rs("��O")=a
rs("���s")=b
rs("����")=c
rs("��T��")=d
rs("���O")=e
rs("�Ȩ�")=wupinll
rs("����")=say
rs.update
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','���~�W�[')"
conn.Execute(sql)
rs.close
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>
<%else%>
<script language=vbscript>
MsgBox "�ܮw�w�g���o�Ӫ��~�F�A�ЧA�K�[�I�O���a"
location.href = "javascript:history.back()"
</script>
<%end if%> 