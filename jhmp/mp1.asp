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
if instr(id,"�x��")<>0 then
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
my=trim(request("my"))
if my<>info(0) then
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,����,�ʧO FROM �Τ� WHERE id="&info(9),conn
if rs.eof or rs.bof then	
message="�A�D���򤤤H�A���n�b�Ƿo�áI"
else
if rs("����")="" or rs("����")="�L" or rs("����")="�C�L"  then
	if rs("����")<2 then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('���Ť����A�ݭn2�ťH�W�ץi�H�[�J�I');window.close();</script>"
		response.end
	end if
	sex=rs("�ʧO")
	rs.close
	rs.open "SELECT �A�X FROM ���� WHERE ����='" & id & "'",conn
	if rs.eof or rs.bof then
		message="���򤤨S��" & id & "�o�Ӫ���"
	else
		if trim(sex)=trim(rs("�A�X")) or trim(rs("�A�X"))="�Ҧ�" then
			message="�A���\�a�[�J�F" & id & "�o�Ӫ���"
			conn.execute "update �Τ� set ����='" & id & "', ����='�̤l',grade=1 where id="&info(9)
			conn.execute "update ���� set �̤l��=�̤l��+1 where ����='" & id & "'"
			info(5)=id
		else
			message="�o�Ӫ������A�X�A"
		end if
	end if
else
	message="�n�[�J" & id & "�A�ЧA���}�A�{�b������"
end if
end if
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
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i���������j["&info(0)&"]</font><font color=FFD7F4>���\�a�[�J�F[" & id & "]�o�Ӫ���</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<title>�F�C���� </title>
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #0000FF; font-family: "�s�ө���"; font-size: 9pt; text-decoration: none }
a:hover      { color: #0000FF; font-family: "�s�ө���"; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" background="../linjianww/0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#FF0000" size="+3">�[�J�������</font></p>
<p align="center"><font color="#FFFFFF"><b><font color="#000000">�F�C����i�ܡG<%=message%>
</font></b></font><font color="#000000"></font></p>
<hr>

</body>

</html>
