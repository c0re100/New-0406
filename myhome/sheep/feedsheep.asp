<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
%>
<!--#include file="data2.asp"-->
<html>
<head>
<title>���i�d��</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" text="#000000" leftmargin="0" topmargin="0">
<table border="1" width="303" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000">
<tr>
<td width="371" height="21">&nbsp;</td>
</tr>
<tr>
<td width="371" height="137">
<%
rs.open"select �W�r,����,�d������ from �d�����A where user='"&info(0)&"'",conn,1,1
if rs.bof then
rs.close
response.write "�z�٨S���ʪ��O�I�֥h�d���ө��R�@���a"%>
<a href="indexsheep.asp">�d���ө�</a>
<%else
%>
<form method="POST" action="checksheep.asp?nick=<%=rs("�W�r")%>">
<table border="0" width="227" cellspacing="1"
cellpadding="0" height="89" align="center">
<tr>
            <td class="p2" width="100%" height="60"> 
              <div align="center"><font color="#0000FF">��-�A�d�����W�r�G<%=rs("�W�r")%> 
                <img src="image/<%=rs("����")%>.gif" border="0" alt="<%=rs("�d������")%> " align="absmiddle"></font><font color="#FFCC00"> 
                </font> </div>
</td>
</tr>
<tr>
<td class="p3" width="100%" height="27">
<p align="center">
<input type="submit"
value="�i�J���i�Ҧ�" name="B1"
style="font-family: �s�ө���; font-size: 12px">
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<%
rs.close
end if
conn.close
%>
</body>

</html>