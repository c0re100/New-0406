<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
%>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select ���� from �d�����A where user='"&info(0)&"' and ��O>0")
if not rs.eof then
money=rs("����")
money=money/2
rs.close
conn.execute("delete from �d�����A where user='"&info(0)&"'")
conn.close
%>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select �Ȩ� from �Τ� where �m�W='"&info(0)&"'")
money1=rs("�Ȩ�")+money
conn.execute("update �Τ� set �Ȩ�='"&money1&"'  where �m�W='"&info(0)&"'")%>
<link rel="stylesheet" href="setup.css">
<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">�A�b����X�F�d���o��<%=money%>��I <%else%> �A�S���d�������X <%rs.close%> <%end if%><br>
<br>
<a href="indexsheep.asp">��^</a></div>