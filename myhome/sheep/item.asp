<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data2.asp"-->
<%
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
sql="SELECT 名字 FROM 寵物狀態 where user='"&info(0)&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
conn.close
set rs=nothing
set conn=nothing
%>
<script language="Vbscript">
msgbox"您的寵物已經死了或您不是這隻寵物的主人！",0,"提示"
history.back
</script>
<%
else
%>
<html>
<head>
<title>使用道具</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" background="../../linjianww/0064.jpg">
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
    <tr> 
      <td width="90" height="22"><font color="#0000FF"><%=rs("道具名字")%></font></td>
      <td width="222" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1">攻擊：<%=rs("攻擊")%> 
          防御：<%=rs("防御")%> 體力：<%=rs("體力")%> </span></font></div>
</td>
      <td width="85" height="22"> 
        <div align="center"><span style="letter-spacing: 1"><%=rs("價格")%>兩</span></div>
</td>
      <td width="59" height="22"> 
        <p align="center"><span style="letter-spacing: 1"><a href="useritem.asp?name=<%=rs("id")%>">使用</a> 
          </span></p>
</td>
</tr>
<%
rs.movenext
loop
conn.close
set rs=nothing
set conn=nothing
end if
%>
</table>

<br>
<font color="#000000">&gt;&gt;<a href="feedsheep.asp">返回寵物育成模式 </a></font></div>
</body>

</html>