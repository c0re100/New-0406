<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data.asp"-->
<%
sql="select * from �^���� where �֦���='" & info(0) & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "�A�٨S���q�O!"
  window.close()
</script>
<%
response.end
end if
%>
<html>
<head>
<title><%=info(0)%>���q�ۤ@��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFFFFF"><b><%=session("myname")%>�A�{���q�۱��p</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td height="18"><font color="#CCCCFF"><b>�q��</b></font></td>
    <td height="18"><font color="#CCCCFF"><b>�ƶq</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">��</font></b></td>
    <td><font color="#FFFFFF"><%=rs("��")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">��</font></b></td>
    <td><font color="#FFFFFF"><%=rs("��")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">��</font></b></td>
    <td><font color="#FFFFFF"><%=rs("��")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="36"> 
      <input type="submit" name="Submit" value="����" onClick="window.close()">
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close
       set rs=nothing
       conn.close
       set conn=nothing%>