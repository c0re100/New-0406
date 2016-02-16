<%@ LANGUAGE="VBScript" %>
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
set rs=server.createobject("adodb.recordset")
sql="delete from 物品買 where id="&request("id")
rs.open sql,conn,1,1
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','物品刪除')"
conn.Execute(sql)
set rs=nothing
conn.close
Response.Redirect "../ok.asp?id=700"
%>
 