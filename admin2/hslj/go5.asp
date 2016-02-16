<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("iq")<>"opy" then
response.redirect "index.asp"
end if
%>
<%
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 攻擊 FROM 用戶 where id="&info(9)
Set Rs=conn.Execute(sql)
at=rs("攻擊")
randomize timer
r=int(rnd*1)
s=int(rnd*5)
at=at*s
randomize timer
r=int(rnd*1)
s=int(rnd*5)
at1=100000*s
if at>at1 then
Session("iq")="ok"
msg="經過一番苦鬥後，你以"&at&"點攻擊贏了打不死的"&at1&"點攻擊，<a href='go6.asp'>按這裡闖下一關吧。</a>"
else
Session("iq")=""
msg="經過一番苦鬥後，你以"&at&"點攻擊不敵打不死的"&at1&"點攻擊，<a href='javascript:self.close()'>按這裡回去吧吧。</a>"
end if
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title>比武結果</title>
</head>
<body bgcolor="#0099FF" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<div align="center"><b><%=msg%></b></div></body>
</html>