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
wa=trim(request.form("wa"))
hy=trim(Request.Form("hy"))
lianbao=trim(Request.Form("lianbao"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 集合 where id=1",conn
conn.Execute "update 集合 set wabao= " & wa& ",hysee=" & hy & ",lianbao="&lianbao&" where id=1"

sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','雜項修改')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
