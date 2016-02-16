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
<!--#include file="homecheck.asp"-->
<!--#include file="data1.asp"-->
<%
rs.open"select restdate from userinfo where user='"&info(0)&"'",conn,1,1
if datediff("d",date,rs("restdate"))=0 then
rs.close
conn.close
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title></head>
<body background="../ljimage/11.jpg">
今天已經休息了一次，別總是閑著呀，找點事情做做吧！：） <a href="#" onclick="history.back();">返回</a> 
<%
else%>
<!--#include file="data3.asp"-->
<%rs.open"select 體力 from 用戶 where 姓名='"&info(0)&"'",conn,1,1

energy=rs("體力")
rs.close
conn.close%>
<!--#include file="data1.asp"-->
<%rs.open"select house from userinfo where user='"&info(0)&"'",conn,1,1

housetype=rs("house")
rs.close
rs.open"select energy from rest where housetype='"&housetype&"'",conn,1,1

energytemp=rs("energy")
energylast=energy+energytemp
tempdate=date
rs.close
conn.close
%>
<!--#include file="data3.asp"-->
<%
conn.execute"update 用戶 set 體力="&energylast&" where 姓名='"&info(0)&"'"
conn.close
%>
<!--#include file="data1.asp"-->
<%
conn.execute"update userinfo set restdate='"&tempdate&"' where user='"&info(0)&"'"
%>
經過休息，您感覺精力恢復了<%=energytemp%>
<a href="#" onclick="history.back();">返回</a>
<%
end if
%>