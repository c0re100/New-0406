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
<title>靈劍江湖</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" TEXT="#000000" BGCOLOR="#FFFFFF" background="../../ljimage/11.jpg">
<p align="center"><FONT SIZE="4">[ 靈劍江湖花店 ]<BR><FONT SIZE="2" COLOR="#993366">鼠標點擊即可購買</FONT></FONT></p><%
rs.open "SELECT id,sm,物品名,銀兩,內力 FROM 物品買 where  類型='鮮花' order by 銀兩",conn
do while not rs.bof and not rs.eof
%> <A HREF="#" ONCLICK="window.open('buyhua.asp?id=<%=rs("id")%>','yuanou','scrollbars=yes,resizable=yes,width=655,height=400')"> 
<img src="../../hcjs/jhjs/001/<%=rs("sm")%>.gif" BORDER="0" ALT="<%=rs("物品名")%> [售價]：<%=rs("銀兩")%>兩 加魅力<%=rs("內力")%>點"></A> 
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
