<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
%>
<html>
<head>
<title>�ޯ�C��</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#FFFFFF" link="#FFFFFF" bgcolor="#003399" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"><br>
�z�d���֦����ޯ�C�� <br>
<span style="letter-spacing: 1"></span> </div>
<div align="center">
<table width="298" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" height="26" bgcolor="#FFFFFF">
<tr bgcolor="#006699">
<td height="10" colspan="4" bgcolor="#0099CC">
<div align="center"><span style="letter-spacing: 1">�� �� �� �d �� �� ��</span></div>
</td>
</tr>
<%
sql="SELECT �ޯ�W�� FROM �ޯ�C�� where �֦���='"&info(0)&"'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor="#666666">
      <td colspan="4" height="12" bgcolor="#0066FF"> 
        <div align="center"><%=rs("�ޯ�W��")%></div>
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
<font color="#000000">[<a href="JavaScript:window.close()">�������f</a>] </font>
</div>
</body>

</html>