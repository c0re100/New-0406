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
if info(0)="" then response.redirect"../../error.asp?id=210"
sql="SELECT �ѧO,����I FROM �d�����A where user='"&info(0)&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z���d���w�g���F�αz���O�o���d�����D�H�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("�ѧO")>=rs("����I") then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���d���֤F�ӷ����F�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
name=request("name")
sql="select ����,���s,��O from �D��C�� where id="&name&" and �֦���='"&info(0)&"'"
set rs=conn.execute(sql)
at=rs("����")
guard=rs("���s")
power=rs("��O")

sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+"&power&",����=����+"&at&",���s=���s+"&guard&" where user='"&info(0)&"'"
conn.Execute(sql)
conn.execute("delete from �D��C�� where id="&name&" and �֦���='"&info(0)&"'")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�D��ϥ�</title>
<link rel="stylesheet" href="setup.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
<table border="1" width="380" cellspacing="1" cellpadding="0" height="173" bordercolor="#000000">
<tr>
<td class="p2" width="100%">
<p align="center"><font size="2">�z�w�g���\�ϥΤF�D�� �A��s�Y�i�ݨ�ĪG�I<br>
<br>
�����W�[<%=at%> ���s�W�[<%=guard%> ��O�W�[<%=power%><br>
<font color="#000000"><br>
</font></font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">
<td class="p3" width="100%" height="19">
<p align="center"><font color="#000000"><a href="feedsheep.asp">&gt;&gt;��^</a>
</font>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>