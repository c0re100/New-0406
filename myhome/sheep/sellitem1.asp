<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
%>
<html>
<head>
<title>你擁有的道具</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" link="#0000FF" vlink="#0000FF" background="../../linjianww/0064.jpg" alink="#0000FF">
<div align="center">你擁有的道具列表<br>
<span style="letter-spacing: 1"><!--#include file="data.asp"--></span> </div>
<div align="center">
  <table width="600" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" height="26" bgcolor="#FFFFFF">
    <tr bgcolor="#FF0000"> 
      <td width="100%" height="10" colspan="4"> 
        <div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">現 
          有 道 具</font></span></div>
</td>
</tr>
    <tr bgcolor="#0033FF"> 
      <td width="90" height="17"> 
        <div align="center"><font color="#FFFFFF">道具名字</font> </div>
</td>
      <td width="222" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">道具效果(增加或減少)</span></font>
</div>
</td>
      <td width="85" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">售
價</span></font> </div>
</td>
      <td width="59" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">操
作</span></font> </div>
</td>
</tr>
<%
sql="SELECT 道具名字,攻擊,防御,體力,價格,id FROM 道具列表 where 擁有者='"& info(0) &"'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr bgcolor="#FFFFFF"> 
      <td width="90" height="22"><font color="#0000FF"><%=rs("道具名字")%></font></td>
      <td width="222" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1">攻擊：<%=rs("攻擊")%>
防御：<%=rs("防御")%> 體力：<%=rs("體力")%> </span></font></div>
</td>
      <td width="85" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1"><%=rs("價格")%>兩</span></font></div>
</td>
      <td width="59" height="22"> 
        <p align="center"><font color="#0000FF"><span style="letter-spacing: 1"><a href="sellitem.asp?name=<%=rs("id")%>">賣出</a>
</span></font></p>
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
</table>

<br>
<font color="#000000">&gt;&gt;<a href="itemshop.asp">返回道具商店</a></font></div>
</body>

</html>