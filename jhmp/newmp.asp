<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if InStr(request.form("subject"),"oR")<>0 or InStr(request.form("subject"),"�Ư���")<>0 or InStr(request.form("subject"),"����")<>0 or InStr(request.form("subject"),"��")<>0 or InStr(request.form("subject"),"�x")<>0 or InStr(request.form("subject"),"�����`��")<>0 or InStr(request.form("subject"),"�x��")<>0 or InStr(request.form("subject"),"Or")<>0 or inStr(request.form("subject"),"&")<>0 or inStr(request.form("subject"),"&&")<>0 or InStr(request.form("subject"),"OR")<>0 or InStr(request.form("subject"),"or")<>0 or InStr(request.form("subject"),"=")<>0 or InStr(request.form("subject"),"`")<>0 or InStr(request.form("subject"),"'")<>0 or InStr(request.form("subject")," ")<>0 or InStr(request.form("subject")," ")<>0 or InStr(request.form("subject"),"'")<>0 or InStr(request.form("subject"),chr(34))<>0 or InStr(request.form("subject"),"\")<>0 or InStr(request.form("subject"),",")<>0 or InStr(request.form("subject"),"<")<>0 or InStr(request.form("subject"),">")<>0 then Response.Redirect "../error.asp?id=54"
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �x�� from ���� where ����='"&trim(request.form("subject"))&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�o�Ӫ����w�g�s�b�F�A�A����ЫجۦP�������I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select ����,�Ȩ�,���� from �Τ� where id="&info(9),conn
if rs("����")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x������԰����Ťj��100�šI');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("�Ȩ�")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x������Ȩ�j��2���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("����")<10000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x��������j��10000�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select �x�� from ���� where �x��='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�O�x���F,�^�쭺�����i�a�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
str="Insert Into ���� (����,²��,�x��,�A�X,�������,�Х߮ɶ�) Values('"
str=str & request.form("subject") & "','"
str=str & request.form("comment") & "','"&info(0)&"','"&trim(request.form("ljxb"))&"',100000000,now())"
conn.execute(str)
conn.execute("update �Τ� set ����='"&trim(request.form("subject"))&"',����='�x��',grade=5 where id="&info(9))
mp=trim(request.form("subject"))
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" background="../linjianww/0064.jpg">
<div align="center">
  <p>&nbsp;</p>
</div>

</body>
</html>
<%
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)=info(0)
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�s�����ϥ͡j"& info(0) &"</font>�j�n�ť�<font color=FFD7F4> "&mp&" </font>�s�����������ߡI</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('�s�����Ыئ��\�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End

%>