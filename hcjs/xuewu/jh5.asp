<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,�ƶq FROM ���~ where �֦���=" & info(9) & " and �ƶq>0 and ���~�W='���s���y'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�S���Ӻد��y!');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
id=rs("id")
sl=rs("�ƶq")
sl1=sl*5
conn.execute "update ���~ set �ƶq=�ƶq-"&sl&" where id="&id
conn.execute "update �Τ� set allvalue=allvalue+"&sl1&" where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�@���Ӻد��y"&sl&"��,�����g���"&sl1&"�I�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
%>
