<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會員費 from 用戶 where id="&info(9),conn
hyfei=rs("會員費")
rs.close
rs.open "select 物品名,說明,銀兩,ID from 物品買 where 類型='卡片' and 銀兩>=1 and 銀兩<5000 order by 銀兩"
%>
<html>
<head>
<title>會員卡片</title>
<link rel="stylesheet" href="../chat/ccss2.css">
</head>
<body bgcolor=#FFFFFF background="../ff.jpg" text="#000000">
<BR>
<div align="center">
<font color=blue><%=info(0)%></font>現在有會員費:<font color=red><b><%=hyfei%></b></font>元整
</div>
<table border=1 align=center width=620 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#3399FF">
<tr>
<td width="10%" height="11"><div align="center">卡片名稱</div></td>
<td width="69%" height="11"><div align="center">功 能</div></td>
<td width="11%" height="11"><div align="center">會費</div></td>
<td height="11" width="10%"><div align="center">操 作</div></td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';">
<td><div align="center"><%=rs("物品名")%> </div></td>
<td><%=rs("說明")%></td>
<td><div align="center"><%=rs("銀兩")%>元</div></td>
<td align="center"><a href="buycard.asp?id=<%=rs("id")%>">購買</a></td>
</tr>
<%
rs.movenext
loop
%>
</table><br><br>
<table border=1 align=center width=620 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
<%rs.close
rs.open "select 物品名,說明,攻擊,ID from 物品買 where 類型='卡片' and 攻擊>=1 order by 銀兩"
do while not rs.bof and not rs.eof
%>
<tr  bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';">
<td width='12%'><div align="center"><%=rs("物品名")%> </div></td>
<td width='65%'><%=rs("說明")%></td></td>
<td width='16%'><div align="center"><%=rs("攻擊")%>個金幣</div></td>
<td align="center" width="10%"><a href="buycard.asp?id=<%=rs("id")%>">購買</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>