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
<title>�d���ө���</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../css.css">
</head>
<body text="#000000" link="#0000FF" vlink="#0000FF" background="../../images/8.jpg" alink="#0000FF">
<div align="center"><font color="#000000">�i<a href="indexsheep.asp">�d���ө�</a>�j</font> 
  <font color="#000000">�i<a href="stunt.asp">�ޯ�ө�</a>�j</font> <font color="#000000">�i<a href="itemshop.asp">�D��ө�</a>�j</font> 
  <font color="#000000"><br>
</font><br>
<br>
</div>
<div align="center">�w����{�d���ө��A�o�̦��U�����P�������d�ѧA���ʨ��C�`�N�A�d������R�@�����C<br>
<span style="letter-spacing: 1"><!--#include file="data.asp"--></span>
<table width="572" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" height="26" bgcolor="#FFFFFF">
    <tr> 
      <td width="100%" height="10" colspan="4" bgcolor="#0099CC"> 
        <div align="center"><span style="letter-spacing: 1">�{ �� �d ��</span></div>
</td>
</tr>
    <tr> 
      <td width="90" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">�d������</span></font>
</div>
</td>
      <td width="222" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">�d���Ѽ�</span></font>
</div>
</td>
      <td width="85" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">��
��</span></font> </div>
</td>
      <td width="59" height="17" bgcolor="#666666"> 
        <div align="center"><font color="#FFFFFF"><span style="letter-spacing: 1">��
�@</span></font> </div>
</td>
</tr>
<%
sql="SELECT �d������,����,���s,��O,���� FROM �d��'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
    <tr> 
      <td width="90" height="22"><font color="#0000FF"><%=rs("�d������")%></font></td>
      <td width="222" height="22"> 
        <div align="center"><font color="#0000FF"><span style="letter-spacing: 1">�����G<%=rs("����")%> 
          ���s�G<%=rs("���s")%> ��O�G<%=rs("��O")%> </span></font></div>
</td>
      <td width="85" height="22"> 
        <div align="center"><span style="letter-spacing: 1"><%=rs("����")%>��</span></div>
</td>
      <td width="59" height="22"> 
        <p align="center"><span style="letter-spacing: 1"><font color="#0080FF"><a href="buy.asp?name=<%=rs("�d������")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></font></span></p>
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