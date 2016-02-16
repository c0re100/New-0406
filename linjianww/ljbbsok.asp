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
leixin=trim(request.form("leixin"))
name=trim(request.form("name"))
wz1=trim(request.form("wz1"))
id=request("id")
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../bbs/bbs.asa")
Conn.Open connstr
if a="m" then
sql="select cname from board where bid>=1"
set rs=Conn.execute(sql)
	nameid=int(abs(request("id")))
	conn.Execute "update board set cname='"&leixin&"',bdinfo='" & name & "',版主='" & wz1 &"' where bid="&nameid
end if
if Request.Form("submit")="刪除" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "delete * from board where bid="&nameid
end if
if Request.Form("submit")="添加" then
sql="select cname from board "
set rs=Conn.execute(sql)
conn.Execute "insert into board(cname,bdinfo,版主) values ('" & leixin & "','" & name & "','" & wz1 & "')"
end if
if a="p" then

jingyan=trim(request.form("jingyan"))
sql="select name from user_info "
set rs=Conn.execute(sql)
conn.Execute "update user_info set name='"&jingyan&"'"
end if
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
