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
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" background="0064.jpg">
<p align="center">&nbsp;  <center> 
  <table border=1 align=center width=550 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor="#FF6600"> 
      <td colspan="4"> 
        <p align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">�{���d�� 
          <A HREF="addcard.asp"><FONT COLOR="#99FFCC" SIZE="2">�s�W�d��</FONT></A> 
          </font></span></p>
      </td></tr> 
<tr bgcolor=#C4DEFF> <td width="117"><font color="#000000"><span style="letter-spacing: 1">�d���W</span></font></td><td width="224"><font color="#000000"><span style="letter-spacing: 1">�\��</span></font></td>
      <td width="101"><font color="#000000"><span style="letter-spacing: 1">���</span></font></td>
      <td width="98"><font color="#000000"><span style="letter-spacing: 1">�ާ@</span></font></td>
    </tr> 
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM ���~�R where  ����='�d��' and �Ȩ�>=1"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%> 
    <tr  bgcolor="#FFFFFF"  onmouseout="this.bgColor='#ffffff';"onmouseover="this.bgColor='#99CCFF';"> 
      <td width="117"><font color="#FF6600" size="-1"><span style="letter-spacing: 1"><%=rs("���~�W")%> 
        </span> </font> 
      <td width="224"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("����")%></font></span></td>
      <td width="101"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("�Ȩ�")%>��</font></span></td>
      <td width="98"> 
        <div align="center"><font size="-1"><span style="letter-spacing: 1"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">�ק�</a> 
          <a href="del.asp?id=<%=rs("id")%>">�R��</a> </span></font></div>
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
sql="SELECT * FROM ���~�R where  ����='�d��' and ����>=1"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr  bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#ffffff';"onMouseOver="this.bgColor='#99CCFF';"> 
      <td width="117"><font color="#FF6600" size="-1"><span style="letter-spacing: 1"><%=rs("���~�W")%> 
        </span> </font> 
      <td width="224"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("����")%></font></span></td>
      <td width="102"><span style="letter-spacing: 1"><font color="#000000" size="-1"><%=rs("����")%>�Ӫ���</font></span></td>
      <td width="97"> 
        <div align="center"><font size="-1"><span style="letter-spacing: 1"><a href="modifyyao.asp?wupinid=<%=rs("id")%>">�ק�</a> 
          <a href="del.asp?id=<%=rs("id")%>">�R��</a> </span></font></div>
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
