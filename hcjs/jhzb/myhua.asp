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
<title> 艶糃打</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="bg.gif">
<p align="center">[ 艶糃打<font color=blue> <%=info(0)%> </font>翧 ]</p>
<%
rs.open "SELECT sm,珇,计秖 FROM 珇 where  摸='翧' and 局Τ="&info(9)&" order by 蝗ㄢ",conn
Response.Write "<div align='center'>"
if rs.eof or rs.bof then Response.Write "<font size=+2 color=red>癸ぃ癬眤⊿Τヴ翧</font>"
do while not rs.bof and not rs.eof
%>
<img src="../../hua/flower<%=rs("sm")%>.gif" alt="<%=rs("珇")%>&nbsp;&nbsp;<%=rs("计秖")%>"> 
<%x=x+1
if x/5=int(x/5) then Response.Write "<br>"
rs.movenext
loop
Response.Write "<br><br>"
Response.Write "公夹氨璶盢穦Τ弧"
Response.Write "</div>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
