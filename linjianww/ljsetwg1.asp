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
wg1=trim(request.form("wg1"))
wg2=trim(request.form("wg2"))
wg3=trim(request.form("wg3"))
wg4=trim(request.form("wg4"))
mp=trim(request.form("mp"))
nl=trim(request.form("nl"))
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT id FROM �Z�\ where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�藍�_�A�S���Ӫ����Z�\�I"
	response.end
end if
rs.close
	nameid=int(abs(request("id")))
	conn.Execute "update �Z�\ set �Z�\='" & wg1 & "', ����='" & wg2 & "',�ŧO=" & wg3 & ",�׷ү�=" & wg4 & ",���O=" & nl & " where id="&nameid
end if
if Request.Form("submit")="�R��" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "delete * from �Z�\ where id="&nameid
end if
if a="n" then
conn.Execute "insert into �Z�\(����,�Z�\,����,�ŧO,�׷ү�,���O) values ('" & mp & "','" & wg1 & "','" & wg2 & "'," & wg3 & "," & wg4 & "," & nl & ")"
end if

sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�Z�\�ק�ާ@')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
