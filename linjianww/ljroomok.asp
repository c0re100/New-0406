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
b=trim(request.form("b"))
c=trim(request.form("c"))
d=trim(request.form("d"))
e=trim(request.form("e"))
f=trim(request.form("f"))
g=trim(request.form("g"))
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT a FROM room where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "對不起，沒有該房間！"
	response.end
end if

	nameid=int(abs(request("id")))
	conn.Execute "update room set a='" & aa & "', b=" & b & ", c=" & c & ", d='" & d & "', e='" & e & "', f=" & f & ", g=" & g & " where id="&nameid
end if
if Request.Form("submit")="刪除" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))

	conn.Execute "delete * from room where id="&nameid
end if
if Request.Form("submit")="添加" then
sql="select * from room "
set rs=Conn.execute(sql)
conn.Execute "insert into room(a,b,c,d,e,f,g) values ('" & aa & "'," & b & "," & c & ",'" & d & "','" & e & "'," & f & "," & g & ")"
end if
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','房間修改操作')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
