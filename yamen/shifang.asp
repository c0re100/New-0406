<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� from �Τ� where id="&info(9),conn
if rs("�Ȩ�")<1000000 then
response.write my & "�A���Ȩ⤣��"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select ���A from �Τ� where �m�W='" & name & "'",conn
if rs("���A")="�ʸT" then
	response.write name & "�Τ�O�Q�ʸT����������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000000 where id="&info(9)
conn.execute "update �Τ� set ���A='���`',�n��=now() where �m�W='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "listlao.asp"
%>
 