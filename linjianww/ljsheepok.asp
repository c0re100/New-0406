<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=request("a")
aa=trim(request.form("a"))
a1=trim(request.form("a1"))
b=trim(request.form("b"))
c=trim(request.form("c"))
d=trim(request.form("d"))
e=trim(request.form("e"))
f=trim(request.form("f"))
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT id FROM �d�� where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�藍�_�A�S�����d���I"
	response.end
end if

	nameid=int(abs(request("id")))
	conn.Execute "update �d�� set  �d������='" & aa & "', �ޯ�='"&a1&"',����='" & b & "',����=" & c & ", ���s=" & d & ",��O=" & e & ", ����=" & f & " where id="&nameid
end if
if Request.Form("submit")="�R��" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))

	conn.Execute "delete * from �d�� where id="&nameid
end if
if Request.Form("submit")="�K�[" then
sql="select id from �d�� "
set rs=Conn.execute(sql)
conn.Execute "insert into �d��(�d������,����,����,���s,��O,����) values ('" & aa & "','"&a1&"','" & b & "'," & c & "," & d & "," & e & "," & f & ")"
end if
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�d���ק�ާ@')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
