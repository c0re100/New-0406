<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><!--#include file="data.asp"--><%
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
else
name=request("name")
if name="���q����" then 
money1=50000 
end if 
if name="�@������" then 
money1=100000 
end if 
if name="������" then 
money1=300000 
end if 
if name="�t�w����" then 
money1=500000 
end if
rs.open"select �Ȩ� from �Τ� where �m�W='"&info(0)&"'",conn,1,1
if rs("�Ȩ�")-money1<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������R�o�ӧޯ�I",0,"FLUSH"
history.back
</script>
<%
else
money=rs("�Ȩ�")-money1
conn.execute("update �Τ� set �Ȩ�='"&money&"'  where �m�W='"&info(0)&"'")
rs.close
%><!--#include file="data2.asp"--><%
set rs=server.createobject("adodb.recordset")
rs.open "select �ޯ�W�� from �ޯ�C�� where �֦���='"&info(0)&"' and �ޯ�W��='"&name&"'",conn,1,1
if not rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�藍�_�A�z�w�g�R�L�F�o�ӧޯ�I');location.href = 'stunt.asp';}</script>"
response.end
end if
rs.close
set rs=server.createobject("adodb.recordset")
sql="select �ޯ�W��,�֦��� from �ޯ�C�� where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("�ޯ�W��")=name
rs("�֦���")=info(0)
rs.update
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�ʶR���\</title>
<link rel="stylesheet" href="setup.css">
</head>

<body bgcolor="#3a4b91" background="../../chat/bk.jpg" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
    <table border="1" width="338" cellspacing="1" cellpadding="0" height="104" bordercolor="#FFFFFF">
      <tr>
<td class="p2" width="100%">
<p align="center"><font size="2">�z�w�g���\�ʶR�F<%=name%>�C<br>
<br>
<br>
<br>
</font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">

<td class="p3" width="100%" height="19">
<p align="center">&gt;&gt;<a href="itemshop.asp"><font color="#FFFFFF">��^�D��ө�</font></a>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>
</html>