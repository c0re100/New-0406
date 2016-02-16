<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Response.Expires=0
Response.Buffer=true
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
<body text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" background="0064.jpg">
<p align="center">&nbsp;  <center> 
  <table border=1 align=center width=550 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor="#FF6600"> 
      <td colspan="4"> 
        <p align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">現有卡片 
          <A HREF="addcard.asp"><FONT COLOR="#99FFCC" SIZE="2">新增卡片</FONT></A> 
          </font></span></p>
      </td></tr> 
<tr bgcolor=#C4DEFF> <td width="117"><font color="#000000"><span style="letter-spacing: 1">卡片名</span></font></td><td width="224"><font color="#000000"><span style="letter-spacing: 1">功能</span></font></td>
      <td width="101"><font color="#000000"><span style="letter-spacing: 1">售價</span></font></td>
      <td width="98"><font color="#000000"><span style="letter-spacing: 1">操作</span></font></td>
    </tr> 
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 where  類型='卡片' and 銀兩>=1"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%> 
    <tr  bgcolor="#FFFFFF"  onmouseout="this.bgColor='#ffffff';"onmouseover="this.bgColor='#99CCFF';"> 
      <td width="117"><font color="#FF6600" size="-1"><span style="letter-spacing: 1"><%=rs("物品名")%> 
        </span> </font> 
      <td width="224"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("說明")%></font></span></td>
      <td width="101"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("銀兩")%>兩</font></span></td>
      <td width="98"> 
        <div align="center"><font size="-1"><span style="letter-spacing: 1"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">修改</a> 
          <a href="del.asp?id=<%=rs("id")%>">刪除</a> </span></font></div>
      </td></tr> 
<%
rs.movenext
loop
%> </table>
  <table border=1 align=center width=550 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 where  類型='卡片' and 攻擊>=1"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr  bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#ffffff';"onMouseOver="this.bgColor='#99CCFF';"> 
      <td width="117"><font color="#FF6600" size="-1"><span style="letter-spacing: 1"><%=rs("物品名")%> 
        </span> </font> 
      <td width="224"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("說明")%></font></span></td>
      <td width="102"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("攻擊")%>個金幣</font></span></td>
      <td width="97"> 
        <div align="center"><font size="-1"><span style="letter-spacing: 1"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">修改</a> 
          <a href="del.asp?id=<%=rs("id")%>">刪除</a> </span></font></div>
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
