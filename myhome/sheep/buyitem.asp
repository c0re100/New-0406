<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
%><!--#include file="data.asp"--><%
name=request("name")
rs.open"select ����,����,���s,��O from �d���D�� where �D��W�r='"&name&"'",conn,1,1
money1=rs("����")
at=rs("����")
guard=rs("���s")
power=rs("��O")
rs.close
rs.open"select �Ȩ�,�ާ@�ɶ� from �Τ� where �m�W='"&info(0)&"'",conn,1,1
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'itemshop.asp';}</script>"
	Response.End
end if
if rs("�Ȩ�")-money1<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������R�o�ӹD��I",0,"FLUSH"
history.back
</script>
<%
else
money=rs("�Ȩ�")-money1
conn.execute("update �Τ� set �Ȩ�="&money&",�ާ@�ɶ�=now()  where �m�W='"&info(0)&"'")
rs.close
%><!--#include file="data2.asp"--><%
set rs=server.createobject("adodb.recordset")
sql="select �D��W�r,����,��O,����,���s,�֦��� from �D��C�� where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("�D��W�r")=name
rs("����")=at
rs("��O")=power
rs("����")=money1
rs("���s")=guard
rs("�֦���")=info(0)
rs.update
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
    <table border="1" width="338" cellspacing="1" cellpadding="0" height="104" bordercolor="#FFFF00">
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