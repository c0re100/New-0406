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
rs.open "SELECT �m�W1 FROM ��� where �m�W1='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=481"
	response.end
end if
rs.close
rs.open "select �t��,�ʧO from �Τ� where �m�W='"&name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=427"
	response.end
else
if rs("�t��")="�L" then
sex=rs("�ʧO")
rs.close
rs.open "SELECT �ʧO,�t�� FROM �Τ� where id=" & info(9),conn
if rs("�ʧO")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=428"
	response.end
end if			
if rs("�t��")="�L" then
	conn.execute "update �Τ� set �t��='" & name & "',���B�ɶ�=date() where id="&info(9)
	conn.execute "update �Τ� set �t��='" & my & "',���B�ɶ�=date() where �m�W='" & name & "'"
	conn.execute "delete * from ��� where �m�W1='" & name & "' or �m�W2='" & my & "'or �m�W1='" & my & "'or �m�W2='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "yuelao.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Redirect "../error.asp?id=429"
response.end
end if
else
conn.close
Response.Redirect "../error.asp?id=430"
response.end
end if
end if
%> 