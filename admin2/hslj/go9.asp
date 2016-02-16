<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache" 
if Session("iq")<>"ok" then
response.redirect "index.asp"
end if
my=ljjh_nickname
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
sql="SELECT 內力 FROM 用戶 where 姓名='" & my & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
Response.Write "<script language=JavaScript>{alert('你不是靈劍江湖中人！');location.href = 'index.asp';}</script>"
set rs=nothing
conn.close
set conn=nothing
Response.End
else
at=rs("內力")
if ljjh_nickname="" then
Response.Write "<script language=JavaScript>{alert('你不是靈劍江湖中人！');location.href = 'index.asp';}</script>"
set rs=nothing
conn.close
set conn=nothing
Response.End
else
Randomize
at=int(at*rnd)
Randomize
at1=int(5000000*rnd)
if at>at1 then
Session("iq")="hla"
msg="經過一番苦鬥後，你以"&at&"點內力贏了林平之的"&at1&"點內力，<a href='go10.asp'>按這裡用內力震開寶藏的門吧。</a>"
else
Session("iq")=""
msg="經過一番苦鬥後，你以"&at&"點內力不敵林平之的"&at1&"點內力，<a href='javascript:self.close()'>按這裡回去吧。</a>"
end if
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title>比武結果</title><LINK href="../style.css" rel=stylesheet></head>
<body background="images/bg.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"><%=msg%></div>
</html>
<%
end if
end if
%>
