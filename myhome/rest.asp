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
<title>�F�C����</title></head>
<body background="../ljimage/11.jpg">
���Ѥw�g�𮧤F�@���A�O�`�O�~�ۧr�A���I�Ʊ������a�I�G�^ <a href="#" onclick="history.back();">��^</a> 
<%
else%>
<!--#include file="data3.asp"-->
<%rs.open"select ��O from �Τ� where �m�W='"&info(0)&"'",conn,1,1

energy=rs("��O")
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
conn.execute"update �Τ� set ��O="&energylast&" where �m�W='"&info(0)&"'"
conn.close
%>
<!--#include file="data1.asp"-->
<%
conn.execute"update userinfo set restdate='"&tempdate&"' where user='"&info(0)&"'"
%>
�g�L�𮧡A�z�Pı��O��_�F<%=energytemp%>
<a href="#" onclick="history.back();">��^</a>
<%
end if
%>