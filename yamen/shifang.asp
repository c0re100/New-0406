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
rs.open "select 銀兩 from 用戶 where id="&info(9),conn
if rs("銀兩")<1000000 then
response.write my & "你的銀兩不夠"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select 狀態 from 用戶 where 姓名='" & name & "'",conn
if rs("狀態")="監禁" then
	response.write name & "用戶是被監禁的不能釋放"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-1000000 where id="&info(9)
conn.execute "update 用戶 set 狀態='正常',登錄=now() where 姓名='" & name & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "listlao.asp"
%>
 