<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><!--#include file="data.asp"--><%
Response.Expires=0
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
else
name=request("name")
rs.open"select ����,����,���s,��O,�d������,�ޯ�,����,id from �d�� where �d������='"&name&"'",conn,1,1
money1=rs("����")
at=rs("����")
guard=rs("���s")
power=rs("��O")
lx=rs("�d������")
jn=rs("�ޯ�")
say=rs("����")
petid=rs("id")
rs.close
rs.open"select �Ȩ� from �Τ� where �m�W='"&info(0)&"'",conn,1,1
if rs("�Ȩ�")-money1<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������R�d���I",0,"FLUSH"
history.back
</script>
<%
else
money=rs("�Ȩ�")-money1
conn.execute("update �Τ� set �Ȩ�='"&money&"'  where �m�W='"&info(0)&"'")
rs.close%>
<!--#include file="data2.asp"-->

<%rs.open"select �W�r from �d�����A where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�@��j�L����R�@���d����",0,"FLUSH"
history.back
</script>
<%
else
set rs=server.createobject("adodb.recordset")
sql="select �W�r,����,�d������,�ޯ�,����,��O,����,���s,����id,user from �d�����A where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("�W�r")=lx
rs("����")=say
rs("�d������")=lx
rs("�ޯ�")=jn
rs("����")=at
rs("��O")=power
rs("����")=money1
rs("���s")=guard
rs("����id")=petid
rs("user")=info(0)
rs.update
rs.close
set rs=nothing	
conn.close
set conn=nothing
end if
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
    <table border="1" width="338" cellspacing="1" cellpadding="0" height="104" bordercolor="#FFFF00">
      <tr>
<td class="p2" width="100%">
<p align="center"><font size="2">�z�w�g���\�ʶR�F<%=lx%> �A�֦^�h���d���_�Ӧnť���W�r�I<br>
<br>
</font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">

<td class="p3" width="100%" height="19">
                <p align="right"><font size="2"><a href="feedsheep.asp" target="_top"><font color="#00FFFF">�ֶ}�l���i<%=lx%>�a�I</font></a></font> 
              </td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>
