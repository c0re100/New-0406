<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%><!--#include file="data2.asp"--><%
sheepname=request.form("sheepname")
if sheepname="" then
response.redirect"Warning.htm"
else
rs.open"select user from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
rs.close
rs.open"select peiplushappy,peipluslife from rules ",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
peiplushappy=rs("peiplushappy")
peipluslife=rs("peipluslife")
rs.close
conn.close%>
<!--#include file="data.asp"--><%
rs.open"select 銀兩 from 用戶 where 姓名='"&info(0)&"'",conn,1,1
tempsplosh=rs("銀兩")
if tempsplosh<0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您沒有足夠的錢！先去賺點錢吧！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
else%>
<!--#include file="data2.asp"-->
<%
rs.open"select feeddate,workload,sheephappy,life from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
feeddate=rs("feeddate")
workload=rs("workload")
sheephappy=rs("sheephappy")+peiplushappy
if sheephappy>100 then
sheephappy=100
end if
life=rs("life")+peipluslife
if life>100 then
life=100
end if
rs.close
if datediff("d",feeddate,date)<>0 then 
tempdate=date
conn.execute"update sheep set life='"&life&"',sheephappy='"&sheephappy&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 銀兩=銀兩-100000 where 姓名='"&info(0)&"'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('陪伴完畢！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
else
if workload>=8 then

	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您已經維護小孩8次！明天再來吧！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 銀兩=銀兩-100000 where 姓名='"&info(0)&"'"
conn.close
%><!--#include file="data1.asp"--><%
conn.execute"update sheep set life='"&life&"',sheephappy='"&sheephappy&"',workload=workload+1 where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('陪伴完畢！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
end if
end if
end if
end if
end if

%>