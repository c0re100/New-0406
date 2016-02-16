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

idd=trim(request.form("id"))
aa=trim(request.form("a"))
bb=trim(Request.Form("b"))
cc=trim(Request.Form("c"))
dd=trim(Request.Form("d"))
ee=trim(Request.Form("e"))
ff=trim(Request.Form("f"))
gg=trim(Request.Form("g"))
hh=trim(Request.Form("h"))
ii=trim(Request.Form("i"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM room where id="&idd,conn
conn.Execute "update room set a='" & aa & "',b=" & bb & ",c=" & cc & ",d='" & dd & "',e='" & ee & "',f=" & ff & ",g=" & gg & ",h=" & hh & ",i=" & ii & " where id="&idd

sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','房間修改操作')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
