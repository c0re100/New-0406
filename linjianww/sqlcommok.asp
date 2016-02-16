<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
sqlstr=Request.Form("sqlstr")
fs=int(Request.form("sqlfs"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Set Rs=conn.Execute(sqlstr)
sqlstr=replace(sqlstr,"'","，")
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：所執行的指定執行完成！點確定反回！');</script>"
%>
<script language="vbscript">
location.href = "javascript:history.go(-1)"
</script>
<%
response.end
%>
