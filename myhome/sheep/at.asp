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
<link rel="stylesheet" href="../../css.css">
</head>
<body leftmargin="0" topmargin="7" bgcolor="#FFFFFF">
<div align="center">
<div align="center">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td valign="top" align="center"> [ <font color="#009900">�d 
�� �� ��</font> ]<br> 
<br>
          �i<a href="indexsheep.asp"> �d���ө� </a>�j �i<a href="stunt.asp"> �ޯ�ө�</a> 
          �j �i <a href="itemshop.asp">�D��ө� </a>�j �i <a href="at.asp">�d���m�Z��</a> �j<br> 
<br> 
<div align="center"> </div> 
<div align="center"><font size="-1">�w����{�d���ө��A�o�̦��U�����P�������d�ѧA���ʨ��C�`�N�A�d������R�@�����C</font><br> 
<!--#include file="data.asp"--> <br> 
<br> 
<table width="572" border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="eeeeee"> 
<tr> 
<td width="100%" colspan="4"> 
<div align="center">�{ �� �d ��</div> 
</td> 
</tr> 
<tr> 
<td width="90"> 
<div align="center">�d������ </div> 
</td> 
<td width="222"> 
<div align="center">�d���Ѽ� </div> 
</td> 
<td width="85"> 
<div align="center">�� �� </div> 
</td> 
<td width="59"> 
<div align="center">�� �@ </div> 
</td> 
</tr> 
<% 
sql="SELECT �d������,����,���s,��O,���� FROM �d��'" 
Set rs=conn.Execute(sql) 
do while not rs.bof and not rs.eof 
%> 
<tr> 
<td width="90"><font color="#0000FF"><%=rs("�d������")%></font></td> 
<td width="222"> 
<div align="center"><font color="#0000FF">�����G<%=rs("����")%> 
���s�G<%=rs("���s")%> ��O�G<%=rs("��O")%> </font></div> 
</td> 
<td width="85"> 
<div align="center"><%=rs("����")%>��</div> 
</td> 
<td width="59"> 
<p align="center"><font color="#0080FF"><a href="buy.asp?name=<%=rs("�d������")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></font></p> 
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
</div> 
</td> 
</tr> 
</table> 
</div> 
</div> 
</body>

</html>