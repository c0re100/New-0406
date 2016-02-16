<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<title></title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<br>
<table border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" width="400">
<tr>
<td height="19" colspan="3">
<div align="center">
<div align="center"><a href="addyao.asp"><font color="#0000ff">添加藥品</font></a></div>
</div>
</td>
<td height="19" width="124">
<div align="center"><a href="yaopu.asp"><font color="#0000ff">補藥</font></a></div>
</td>
<td height="19" colspan="2" width="133">
<div align="center"><a href="yaopu2.asp"><font color="#0000ff">毒藥</font></a></div>
</td>
</tr>
</table>
<p align="center"> &nbsp;
<script src="../hcjs/2.js"></script>
<center>
  <table border="1" align="center" width="600" cellpadding="0"
cellspacing="0">
    <tr bgcolor="#00CCFF"> 
      <td height="17" colspan="6"> 
        <p align="center"><span style="letter-spacing: 1">現 有 藥 品</span></p>
    <tr> 
      <td height="18" bgcolor="#C4DEFF" width="75"> 
        <div align="center"> <font color="#000000"><span style="letter-spacing: 1">藥品名</span></font></div>
      </td>
      <td bgcolor="#C4DEFF" height="18" width="92"> 
        <div align="center"><font color="#000000">圖 片</font> </div>
      </td>
      <td bgcolor="#C4DEFF" height="18" width="148"> 
        <div align="center"><font color="#000000">圖片說明</font></div>
      </td>
      <td bgcolor="#C4DEFF" height="18" width="132"> 
        <div align="center"><font color="#000000">功 能</font></div>
      </td>
      <td height="18" bgcolor="#C4DEFF" width="65"> 
        <div align="center"> 售 價 </div>
      </td>
      <td height="18" bgcolor="#C4DEFF" width="74"> 
        <div align="center"> <font color="#000000">操 作</font> </div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 where  類型='毒藥'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="11" width="75"><%=rs("物品名")%> 
      <td height="11" width="92"> 
        <div align="center"><img src="../hcjs/jhjs/001/<%=rs("說明")%>.gif" border="0" alt="<%=rs("物品名")%> "></div>
      </td>
      <td height="11" width="148">../hcjs/jhjs/001/<%=rs("說明")%>.gif</td>
      <td height="11" width="132"><span style="letter-spacing: 1"><font size="-1">補內力<%=rs("內力")%>，補生命<%=rs("體力")%></font></span></td>
      <td height="11" width="65"><%=rs("銀兩")%>兩</td>
      <td height="11" width="74"> 
        <div align="center"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">修改</a> 
          <a href="del.asp?id=<%=rs("id")%>">刪除</a></div>
      </td>
    </tr>
    <%
rs.movenext
loop
%>
  </table>
</center>
</body>
</html>
