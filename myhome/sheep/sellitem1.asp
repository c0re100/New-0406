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
<title>�A�֦����D��</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" link="#0000FF" vlink="#0000FF" background="../../linjianww/0064.jpg" alink="#0000FF">
<div align="center">�A�֦����D��C��<br>
<span style="letter-spacing: 1"><!--#include file="data.asp"--></span> </div>
<div align="center">
  <table width="600" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" height="26" bgcolor="#FFFFFF">
    <tr bgcolor="#FF0000"> 
      <td width="100%" height="10" colspan="4"> 
        <div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">�{ 
          �� �D ��</font></span></div>
</td>
</tr>
    <tr bgcolor="#0033FF"> 
      <td width="90" height="17"> 
        <div align="center"><font color="#FFFFFF">�D��W�r</font> </div>
</td>
      <td width="222" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">�D��ĪG(�W�[�δ��)</span></font>
</div>
</td>
      <td width="85" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">��
��</span></font> </div>
</td>
      <td width="59" height="17"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">��
�@</span></font> </div>
</td>
</tr>
<%
sql="SELECT �D��W�r,����,���s,��O,����,id FROM �D��C�� where �֦���='"& info(0) &"'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr bgcolor="#FFFFFF"> 
      <td width="90" height="22"><font color="#0000FF"><%=rs("�D��W�r")%></font></td>
      <td width="222" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1">�����G<%=rs("����")%>
���s�G<%=rs("���s")%> ��O�G<%=rs("��O")%> </span></font></div>
</td>
      <td width="85" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1"><%=rs("����")%>��</span></font></div>
</td>
      <td width="59" height="22"> 
        <p align="center"><font color="#0000FF"><span style="letter-spacing: 1"><a href="sellitem.asp?name=<%=rs("id")%>">��X</a>
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
<font color="#000000">&gt;&gt;<a href="itemshop.asp">��^�D��ө�</a></font></div>
</body>

</html>