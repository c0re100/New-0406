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
sql="select * from �^���� where �֦���='" & info(0) & "' and (��>0 or ��>0 or ��>0)"
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
 jin=rs("��")*10
 yin=rs("��")*6
 tong=rs("��")*3
 money=jin+yin+tong
 conn.execute("update �^���� set ��=0,��=0,��=0 where �֦���='" & info(0) & "'")
 set rs=nothing
 conn.close
 set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
 conn.execute("update �Τ� set �Ȩ�=�Ȩ�+" & money & " where id="&info(9))
 conn.close
 set conn=nothing
%>
<html>
<head>
<title>�q�ۦ��ʳ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<br>
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFCCCC"><b><%=info(0)%></b></font><b><font color="#FFFFFF">�A�⨭�W�q�ۥ���F�A��o�Ȩ�</font><font color="#FFCCCC"><%=money%></font><font color="#FFFFFF">��</font></b></td>
  </tr>
  <tr align="center" bgcolor="#376D95">
    <td colspan="2" height="26">
      <input type="submit" name="Submit" value="����" onclick="window.close()">
    </td>
  </tr>
</table>
</body>
</html>
