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
rs.open "select �_��,���� from �Τ� where id="&info(9),conn
if rs("����")<1000  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����C��1000�Ҥ��X���򪺡I');window.close();}</script>"
Response.End
end if
if rs("�_��")<>"�F�C������"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�S���_���F�C�����۰ڡI');window.close();}</script>"
Response.End
end if
zhizi=rs("����")
conn.execute "update �Τ� set �k�O=�k�O+"&zhizi&",����=0,�_��='�L' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���_���\�A�A���k�O�W��"&zhizi&"�I');window.close();}</script>"%>