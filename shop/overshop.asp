<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then  response.redirect "../error.asp?id=425"
%><!--#include file="data.asp"--><%
sql="SELECT * FROM �ө�"
rs1.open sql,conn1,1,1
if rs1.EOF or rs1.BOF then Response.Redirect "../error.asp?id=484"
%>
<head>
<title>�K�[�ө�</title>
</head>

<body bgcolor="#99CCFF" BACKGROUND="../bg.gif" TEXT="#99CC99">
<p align="center"><FONT COLOR="#800080" SIZE="6">�ק�ө�</FONT></p><DIV ALIGN="LEFT"><%
while not rs1.eof 
%></DIV><CENTER> <A HREF="#" ONCLICK="window.open('overshop1.asp?shopname=<%=rs("�ө��W")%>','yuanou','scrollbars=yes,resizable=yes,width=655,height=400')"> 
<%=rs1("�ө��W")%></A>---�ѪO:<%=rs1("���D")%><BR><BR></CENTER><%
rs.MoveNext 
wend
set rs1=nothing
conn1.close
set conn1=nothing
%><DIV ALIGN="CENTER"><A HREF="index.asp">��^</A></DIV>