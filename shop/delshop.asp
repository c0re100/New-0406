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
shopname=rs1("�ө��W")
%>
<head>
<title>�K�[�ө�</title>
</head>

<body bgcolor="#99CCFF" BACKGROUND="../bg.gif" LINK="#FFFF00" TEXT="#993366">
<p align="center">&nbsp;</p><TABLE WIDTH="75%" BORDER="1" ALIGN="CENTER" CELLPADDING="0" CELLSPACING="1" BGCOLOR="#9999FF"><TR><TD HEIGHT="17"><P ALIGN="CENTER"><FONT SIZE="3">�R</FONT><FONT SIZE="3">��</FONT><FONT SIZE="3">��</FONT><FONT SIZE="3">��</FONT></P></TD></TR><TR><TD HEIGHT="141"><P><%
while not rs1.eof 
%></P><CENTER> <%=rs1("�ө��W")%>--�ѪO:<%=rs1("���D")%> <A HREF="del.asp?id=<%=rs1("�ө��W")%>">�R��</A><BR><BR></CENTER><%
rs1.MoveNext 
wend
set rs1=nothing
conn1.close
set conn1=nothing
%><BR><BR><DIV ALIGN="RIGHT"><P>&nbsp;</P></DIV></TD></TR><TR><TD><DIV ALIGN="CENTER"><A HREF="addshop.asp">�K�[�ө�</A></DIV></TD></TR><TR><TD><DIV ALIGN="CENTER"><A HREF="overshop.asp">�ק�ө�</A></DIV></TD></TR></TABLE><DIV ALIGN="CENTER"><P>&nbsp;</P><P><A HREF="index.asp">��^</A></P></DIV><P>&nbsp;</P>