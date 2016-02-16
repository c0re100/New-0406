<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

name=Request.Form ("name")
job=Request.Form ("job")
shopname=Request.Form ("shopname")

sql="select 職業 from 用戶 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language="vbscript">
alert("沒有此人！")
history.back(1)
</script>
<%
end if
if rs("職業")="配藥師" then 
%>
<script language="vbscript">
alert("此人已是配藥師！")
history.back(1)
</script>
<%
end if
sql="update 用戶 set 職業='"& job &"',所屬='"&shopname&"' where 姓名='" & name & "'"
conn.execute sql
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('您把"&name&"招收為你的："& job &"點擊確定返回！');</script>"
%>
<script language="vbscript">
history.back(1)
</script>