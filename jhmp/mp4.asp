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
if info(0)="�����`��" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	conn.execute "update ���� set �̤l��=�̤l��-1 where ����='" & id & "'"
	conn.execute "delete from �Τ� where �m�W='" & you & "'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�����G{"&id&"}���̤l["&you&"]�R�����\�I');location.href = 'javascript:history.back()';}</script>"
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
end if
%>
