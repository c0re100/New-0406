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
<title>�������</title>
<body leftmargin="0" topmargin="0" bgcolor="#660000">
<%if info(0)="" then Response.Redirect "../error.asp?id=210"
my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then
 	Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�Ȩ� from �Τ� where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if trim(rs("����"))="�C�L" or trim(rs("����"))="�L" or trim(rs("����"))="����"  then
	Response.Write "<script Language=Javascript>alert('�A�O�C�L�A�٨S���ۤv�������A�A�Q�d����r�H');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if rs("�Ȩ�")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�����n�������r�I');location.href = 'javascript:history.back()';</script>"
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"& money &",�D�w=�D�w+"&int(money/500)&",�������=�������+"&money &",�ާ@�ɶ�=now() where id="&info(9)
conn.execute "update ���� set �������=�������+"& money &" where ����='" & info(5) &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�z�V�A�����������m�F"& money &"��A�D�w�W���G"&int(money/500)&"�I�������̤l��A�P�E���ɡI');location.href = 'javascript:history.back()';</script>"
%>
