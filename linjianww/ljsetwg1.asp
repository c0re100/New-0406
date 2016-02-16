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
rs.open "SELECT id FROM 武功 where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "對不起，沒有該門派武功！"
	response.end
end if
rs.close
	nameid=int(abs(request("id")))
	conn.Execute "update 武功 set 武功='" & wg1 & "', 類型='" & wg2 & "',級別=" & wg3 & ",修煉級=" & wg4 & ",內力=" & nl & " where id="&nameid
end if
if Request.Form("submit")="刪除" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "delete * from 武功 where id="&nameid
end if
if a="n" then
conn.Execute "insert into 武功(門派,武功,類型,級別,修煉級,內力) values ('" & mp & "','" & wg1 & "','" & wg2 & "'," & wg3 & "," & wg4 & "," & nl & ")"
end if

sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','武功修改操作')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
