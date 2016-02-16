<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>查看物品</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" topmargin="0" leftmargin=0 background=bk.jpg>
<div align="center"><img src="ico/girl01.gif" width="24" height="41"><img src="ico/girl02.gif" width="21" height="38"> 
</div>
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="../hcjs/es/mybxg.asp" target="_blank"><img src="ico/03.gif" width="9" height="11" border="0">江湖保險櫃</a><img src="ico/03.gif" width="9" height="11" border="0"></div></td></tr>
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="../hcjs/jhzb/eat.asp" target="_blank"><img src="ico/10.gif" width="9" height="9" border="0">在線吃藥</a><img src="ico/10.gif" width="9" height="9"></div></td></tr>
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="#" onClick="window.open('../hcjs/jhzb/head.asp','','scrollbars=yes,resizable=yes,width=750,height=452')"><img src="ico/4.gif" width="11" height="11" border="0">武器裝備</a><img src="ico/4.gif" width="11" height="11"></div></td></tr>
<tr align=center bgcolor=ffcc00><td>物 品 名</td><td>數量</td><td>自用</td></tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "delete * from 物品 where 數量<=0"
rs.open "SELECT * FROM 物品 WHERE 擁有者=" & info(9) & " and 裝備=false order by 類型,銀兩 ",conn
do while not rs.bof and not rs.eof
lx=trim(rs("類型"))
%>
<tr bgcolor="#EEFFEE"  onmouseout="this.bgColor='#EEFFEE';"onmouseover="this.bgColor='#DFEFFF';">
<td width="69">
<div align="center"><%=rs("物品名")%> <img src="../hcjs/jhjs/001/<%=rs("sm")%>.gif" border="0" width="32" height="32"%></div>
</td>
<td width="26">
<div align="center"><%=rs("數量")%> </div>
<div align="center"></div>
</td>
<td width="26">
<div align="center">
<%if trim(rs("類型"))="毒藥" then%>
<a href="javascript:s('/下毒$ <%=rs("物品名")%>')" title=體力:<%=rs("體力")%>內力:<%=rs("內力")%>銀兩:<%=rs("銀兩")%>><font color="#FF0000">下毒</font></a>
<font color="#FFFFFF">
<%end if%>
<%if trim(rs("類型"))="暗器" then%>
</font> <a href="javascript:s('/投擲$ <%=rs("物品名")%>')" title=體力:<%=rs("體力")%>內力:<%=rs("內力")%>銀兩:<%=rs("銀兩")%>><font color="#FF0000">投擲</font></a>
<%end if%>
<%if trim(rs("類型"))="卡片" then%>
<a href="javascript:s('/使用$ <%=rs("物品名")%>')" title=說明:<%=rs("說明")%>價錢:<%=rs("銀兩")%>元><font color="#FF0000">使用</font></a>
<a href="javascript:s('/贈送$ <%=rs("物品名")%>&1')" title=說明:<%=rs("說明")%>價錢:<%=rs("銀兩")%>元><font color="#FF0000">贈送</font></a>
<%end if%>
<%if trim(rs("類型"))<>"暗器" and trim(rs("類型"))<>"毒藥" and trim(rs("類型"))<>"卡片" then%>
<a href="javascript:s('/贈送$ <%=rs("物品名")%>&1')" title=攻擊:<%=rs("攻擊")%>防御:<%=rs("防御")%>體力:<%=rs("體力")%>內力:<%=rs("內力")%>堅固度:<%=rs("堅固度")%>銀兩:<%=rs("銀兩")%>說明:<%=rs("說明")%>><font color="#FF0000">贈送</font></a>
<%end if%>
<a href="javascript:s('/丟棄$ <%=rs("物品名")%>&1')" title=攻擊:<%=rs("攻擊")%>防御:<%=rs("防御")%>體力:<%=rs("體力")%>內力:<%=rs("內力")%>銀兩:<%=rs("銀兩")%>說明:<%=rs("說明")%>><font color="#0000FF">扔了</font></a>
</div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table></body></html>
