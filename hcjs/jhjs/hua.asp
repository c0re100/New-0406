<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
%>
<html>
<head>
<title>�F�C����</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" TEXT="#000000" BGCOLOR="#FFFFFF" background="../../ljimage/11.jpg">
<p align="center"><FONT SIZE="4">[ �F�C����ᩱ ]<BR><FONT SIZE="2" COLOR="#993366">�����I���Y�i�ʶR</FONT></FONT></p><%
rs.open "SELECT id,sm,���~�W,�Ȩ�,���O FROM ���~�R where  ����='�A��' order by �Ȩ�",conn
do while not rs.bof and not rs.eof
%> <A HREF="#" ONCLICK="window.open('buyhua.asp?id=<%=rs("id")%>','yuanou','scrollbars=yes,resizable=yes,width=655,height=400')"> 
<img src="../../hcjs/jhjs/001/<%=rs("sm")%>.gif" BORDER="0" ALT="<%=rs("���~�W")%> [���]�G<%=rs("�Ȩ�")%>�� �[�y�O<%=rs("���O")%>�I"></A> 
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%> 
</body>

</html>
