<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=trim(request("id"))
you=trim(request("you"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,���� FROM �Τ� WHERE id="&info(9),conn
if rs("����")="�x��" and rs("����")=id then
	rs.close
    rs.open "SELECT ����,���� FROM �Τ� WHERE �m�W='"&you&"'",conn
 if trim(rs("����"))="�x��"  then
rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		response.end
	 end if	
 if trim(rs("����"))<>id  then
rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		response.end
	 end if
	message="���\����" & you & "�v�X" & rs("����") & "�I"
	conn.execute "update ���� set �̤l��=�̤l��-1 where ����='" & id & "'"
	conn.execute "update �Τ� set ����='�C�L',����='�̤l',grade=1 where �m�W='" & you & "'"
else
	rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		response.end
'message="�A���O�o�Ӫ������x��"
end if

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
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i���������j["&info(0)&"]</font><font color=FFD7F4>���\����[" & you & "]�v�X[" & rs("����") & "]�I</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title> �F�C���� </title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body background="../images/8.jpg" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FF0000" size="+3">�̤l�޲z</font><br>
<br>
����i�ܡG<%=message%> </div>
<hr>
</body>
</html>