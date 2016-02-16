<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>寵物商店區</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../css.css">
</head>
<body text="#000000" link="#0000FF" vlink="#0000FF" background="../../images/8.jpg" alink="#0000FF">
<div align="center"><font color="#000000">【<a href="indexsheep.asp">寵物商店</a>】</font> 
  <font color="#000000">【<a href="stunt.asp">技能商店</a>】</font> <font color="#000000">【<a href="itemshop.asp">道具商店</a>】</font> 
  <font color="#000000"><br>
</font><br>
<br>
</div>
<div align="center">歡迎光臨寵物商店，這裡有各式不同種類的寵供你選購呵。注意，寵物隻能買一隻呵。<br>
<span style="letter-spacing: 1"><!--#include file="data.asp"--></span>
<table width="572" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" height="26" bgcolor="#FFFFFF">
    <tr> 
      <td width="100%" height="10" colspan="4" bgcolor="#0099CC"> 
        <div align="center"><span style="letter-spacing: 1">現 有 寵 物</span></div>
</td>
</tr>
    <tr> 
      <td width="90" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">寵物類型</span></font>
</div>
</td>
      <td width="222" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">寵物參數</span></font>
</div>
</td>
      <td width="85" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">售
價</span></font> </div>
</td>
      <td width="59" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">操
作</span></font> </div>
</td>
</tr>
<%
sql="SELECT 寵物類型,攻擊,防御,體力,價格 FROM 寵物'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr> 
      <td width="90" height="22"><font color="#0000FF"><%=rs("寵物類型")%></font></td>
      <td width="222" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1">攻擊：<%=rs("攻擊")%> 
          防御：<%=rs("防御")%> 體力：<%=rs("體力")%> </span></font></div>
</td>
      <td width="85" height="22"> 
        <div align="center"><span style="letter-spacing: 1"><%=rs("價格")%>兩</span></div>
</td>
      <td width="59" height="22"> 
        <p align="center"><span style="letter-spacing: 1"><font color="#0080FF"><a href="buy.asp?name=<%=rs("寵物類型")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></font></span></p>
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
</body>

</html>