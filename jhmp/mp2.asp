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
rs.open "SELECT ����,����,�s�� FROM �Τ� WHERE id="&info(9),conn
if rs.eof or rs.bof then
	message="�A�D���򤤤H"
else
	if rs("����")="" or rs("����")="�C�L" or rs("����")="�L" then
		message="�A�õL����"
	else
		if rs("����")="�x��" then 
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('���Q�@�F���@�n�A�O���O�Q���I�I');window.close();</script>"
			response.end
		end if
if info(4)=0 then
	message="�A���}�F" & rs("����") & "�A�l�����O1000,�Ȩ�Φs�ڧ�����1000��I"
	if rs("�s��")>1000 then
		conn.execute "update �Τ� set ����='�C�L',����='�̤l', ���O=���O-1000,grade=1,�Ȩ�=1000,�s��=1000 where id="&info(9)
	else
		conn.execute "update �Τ� set ����='�C�L',����='�̤l', ���O=���O-1000,grade=1,�Ȩ�=0 where id="&info(9)
	end if
else
	message="�A���}�F" & rs("����") & "�A�]���A�O����|���A�ҥH�즳���O�A�������ܡI"	
	conn.execute "update �Τ� set ����='�C�L',����='�̤l',grade=1 where id="&info(9)
end if
	conn.execute "update ���� set �̤l��=�̤l��-1 where ����='" & id & "'"
	info(5)=""
end if

end if
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

<body bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"><font size="+3" color="#FF0000">���}����</font><br>
<br>
<br>
�F�C����i�ܡG<%=message%> </div>
</body>
</html>
