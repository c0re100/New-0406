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
rs.open "Select 銀兩 from 用戶 where id="&info(9),conn
ying=rs("銀兩")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>賭骰子</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>

<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../images/8.jpg">
<div align="center"><font size="+3"><font class=c>賭骰子</font></font><br>
</div>
<div align=center>
<p>- 銀子最少下注是 <b><font color="#0000FF">10</font> 兩</b> -<br>
最大下注銀子是 <font color="#CC0000"><b><font color="#0000FF">3000</font></b></font><b>
兩 </b></p>
<form method="POST" action="dicedispose.asp">
<table border=1 cellspacing=0 cellpadding=0 align="center" width="307" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" bordercolorlight="#000000">
<tr bgcolor="#336633">
<td width="337" height="25"><font size="2" class=c><b>&nbsp;&nbsp;請下注</b></font></td>
</tr>
<tr bgcolor=#CCCCCC>
<td align=center width="337"><font color="#008000" size="2">你現在一共有銀子 <b><%=yin%>
兩</b> 可以作為賭注</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337"><font color="#000000" size="2">我要下注：
<input type="text" name="splosh" size="10" value="0">
&nbsp;兩</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337">
<input type="submit" value="下注啦！！！" name="B1" >
<input type="reset" value="我要考慮一下：）" name="B2" >
</td>
</tr>
</table>
</form>

<p align="center">警告：每次賭錢，你的魅力值會扣 <font color="#FF0000">1 </font>！<br>
<br>
</p>
</div>

</BODY>
</HTML>
