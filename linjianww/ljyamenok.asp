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
shenfen=trim(request.form("shenfen"))
shenfen1=trim(request.form("shenfen1"))
name1=trim(request.form("name1"))
id=request("id")
dj=trim(request.form("dj"))
dj1=trim(request.form("dj1"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT 身份 FROM 用戶 where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "對不起，沒有該用戶！"
	response.end
end if
if rs("身份")="正站長" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('此人是正站長，不能修改！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if shenfen="正站長" or dj=11 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('正站長隻能一個不得修改！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if info(2)=10 and dj>=10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "update 用戶 set 身份='" & shenfen & "', grade=" & dj & " where id="&nameid
end if
if Request.Form("submit")="開除" then
	nameid=int(abs(request("id")))

	conn.Execute "update 用戶 set 門派='遊俠',身份='弟子', grade=1 where id="&nameid
end if
if Request.Form("submit")="添加" then
if shenfen1="正站長" or dj1=11 then
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('為防止江湖混亂正站長隻能有一個！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if info(2)=10 and dj1>=10 then
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.open "SELECT id FROM 用戶 where 姓名='"&name1&"'",conn
nameid1=rs("id")
	conn.Execute "update 用戶 set 門派='官府',身份='" & shenfen1 & "', grade=" & dj1 & " where id="&nameid1
end if
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','官府人員修改操作')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
