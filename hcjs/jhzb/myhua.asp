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
<title> �F�C����</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="bg.gif">
<p align="center">[ �F�C����<font color=blue> <%=info(0)%> </font>�A�� ]</p>
<%
rs.open "SELECT sm,���~�W,�ƶq FROM ���~ where  ����='�A��' and �֦���="&info(9)&" order by �Ȩ�",conn
Response.Write "<div align='center'>"
if rs.eof or rs.bof then Response.Write "<font size=+2 color=red>�藍�_�A�z�S�������A��I</font>"
do while not rs.bof and not rs.eof
%>
<img src="../../hua/flower<%=rs("sm")%>.gif" alt="<%=rs("���~�W")%>&nbsp;&nbsp;<%=rs("�ƶq")%>��"> 
<%x=x+1
if x/5=int(x/5) then Response.Write "<br>"
rs.movenext
loop
Response.Write "<br><br>"
Response.Write "���а��n��N�|�������I"
Response.Write "</div>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
