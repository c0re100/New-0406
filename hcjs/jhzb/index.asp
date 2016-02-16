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
name=info(0)
rs.open "SELECT 攻擊加,防御加,體力加,內力加,姓名,銀兩,allvalue,grade,攻擊,防御,內力,等級,體力 FROM 用戶 WHERE id="& info(9),conn
gjj=rs("攻擊加")
fyj=rs("防御加")
tlj=rs("體力加")
nlj=rs("內力加")
%>
<html>
<head>
<title>我的包袱</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#FFFFFF" background="back.gif">
<div align="center">
<table width="509" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="73" rowspan="3" valign="top">
<div align="center">
<input onClick="javascript:window.document.location.href='index.asp'" title=裝備一覽 type=button value="裝備一覽" name="button7">
<br>
          <input onClick="javascript:window.document.location.href='head.asp'" title=武器裝備 type=button value="武器裝備" name="button">
<br>
          <br>
          <input onClick="javascript:window.document.location.href='eat.asp'" title=喫用藥物 type=button value="喫用藥物" name="button62">
</div>
</td>
<td valign="top" width="436">
<div align="center"><img src="dao.gif" width="40" height="15">我的裝備一覽<img src="dao1.gif" width="40" height="15">
<font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font>
</div>
</td>
</tr>
<tr>
<td valign="top" width="436" height="27">
<div align="center">用戶<font color="#FF0000"><%=rs("姓名")%></font>狀態，紅色為上限！<font color="#FF0000"><font color="#000000">銀子：<%=rs("銀兩")%>
經驗：<%=rs("allvalue")%></font></font><font color="#FF0000"><font color="#000000">
等級：<%=rs("grade")%>級</font></font><br>
攻擊力：<%=rs("攻擊")%>
防御力：<%=rs("防御")%>
內力：<%=rs("內力")%><font color="#FF0000">/<%=rs("等級")*64+2000+nlj%></font>
體力：<%=rs("體力")%><font color="#FF0000">/<%=rs("等級")*240+5260+tlj%>
</font></div>
</td>
</tr>
<tr>
<td width="436" valign="top" height="172">
<table width="78%" border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" align="center">
<tr>
<td> 頭盔：
<%
rs.close
rs.open "SELECT 物品名,防御 FROM 物品 WHERE 擁有者="& info(9) & " and 裝備=true and 類型='頭盔'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 防御：<%=rs("防御")%>
<%end if%>
</td>
</tr>
<tr>
<td> 裝飾：
<%
rs.close
rs.open "SELECT 物品名,防御 FROM 物品 WHERE 擁有者=" & info(9) & " and 裝備=true and 類型='裝飾'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 防御：<%=rs("防御")%>
<%end if%>
</td>
</tr>
<tr>
<td> 盔甲：
<%
rs.close
rs.open "SELECT 物品名,防御 FROM 物品 WHERE 擁有者=" & info(9) & " and 裝備=true and 類型='盔甲'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 防御：<%=rs("防御")%>
<%end if%>
</td>
</tr>
<tr>
<td height="2"> 左手：
<%
rs.close
rs.open "SELECT 物品名,防御 FROM 物品 WHERE 擁有者="&info(9)& " and 裝備=true and 類型='左手'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 防御：<%=rs("防御")%>
<%end if%>
</td>
</tr>
<tr>
<td> 右手：
<%
rs.close
rs.open "SELECT 物品名,攻擊 FROM 物品 WHERE 擁有者=" & info(9) & " and 裝備=true and 類型='右手'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 攻擊：<%=rs("攻擊")%>
<%end if%>
</td>
</tr>
<tr>
<td> 雙腳：
<%
rs.close
rs.open "SELECT 物品名,防御 FROM 物品 WHERE 擁有者=" & info(9) & " and 裝備=true and 類型='雙腳'",conn
if rs.eof or rs.bof then
%>
未裝備
<%else%>
<%=rs("物品名")%> 防御：<%=rs("防御")%>
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<table width="475" border="0" cellspacing="0" cellpadding="0">
<tr>
<td valign="top" align="center">
<input onClick="JavaScript:window.close()" title=關閉窗口 type=button value="關閉窗口" name="button2">
</td>
</tr>
</table>
</div>
</body>
</html>
