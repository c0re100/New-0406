<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
pai=info(5)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title><%=pai%>�Z��</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgproperties="fixed" topmargin="0" leftmargin="0" background=bk.jpg>
<div align="center">
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
<tr><td width="53%" height="20" bgcolor="#FF6633">
<div align="center"><font color="#330033" size="2">�ۦ��W</font></div></td>
<td width="24%" height="20" bgcolor="#FF6633"><div align="center">����</div></td>
<td width="23%" height="20" bgcolor="#FF6633"><font color="#330033" size="2">���O</font></td>
</tr>
<tr bgcolor=EEFFEE><td colspan="3" height="20"><div align="center">�����N</div></td></tr>
<%
rs.open "SELECT �Z�\,���O,�ŧO FROM �Z�\ where ����='" & pai & "' and ����='����'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20"><div align="center"><a href="javascript:s('/�o��$ <%=rs("�Z�\")%>')">
<input type=hidden name=wg1 value='<%=rs("�Z�\")%>'><%=rs("�Z�\")%></a></div></td>
<td height="20"><%=rs("�ŧO")%></td><td height="20"><%=rs("���O")%></td>
</tr>
<%rs.movenext
loop
rs.close%>
<tr  bgcolor="#EEFFEE">
<td colspan="3" height="20">
<div align="center">���s�N</div>
</td>
</tr>
<%
rs.open "SELECT �Z�\,���O,�ŧO FROM �Z�\ where ����='" & pai & "' and ����='���s'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20">
<div align="center"><a href="javascript:s('/�o��$ <%=rs("�Z�\")%>')">
<input type=hidden name=wg1 value='<%=rs("�Z�\")%>'>
<%=rs("�Z�\")%></a></div>
</td>
<td height="20"><%=rs("�ŧO")%></td>
<td height="20"><%=rs("���O")%></td>
</tr>
<%
rs.movenext
loop
rs.close%>
<tr  bgcolor="#EEFFEE">
<td colspan="3" height="20">
<div align="center">��_�N</div>
</td>
</tr>
<%rs.open "SELECT �Z�\,���O,�ŧO FROM �Z�\ where ����='" & pai & "' and ����='��_'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20">
<div align="center"><a href="javascript:s('/�o��$ <%=rs("�Z�\")%>')">
<input type=hidden name=wg1 value='<%=rs("�Z�\")%>'>
<%=rs("�Z�\")%></a></div>
</td>
<td height="20"><%=rs("�ŧO")%></td>
<td height="20"><%=rs("���O")%></td>
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
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >

<tr>
<td colspan="3" height="17" bgcolor="#FF6633">
<div align="center"><font color="#FFFFFF">
<input type="checkbox" name="liji" value="on">
<font color="#000000">�ߧY����</font> </font></div>
</td>
</tr>
</table>
</div>
<p class=p1 align="center"><font color="#FFFFFF" size="2">���ˤO����<br>
���H�]����*���O+�Z�\+�����O�^<br>
- <br>
���]����*���O+�Z�\+���s�O�^</font></p>
</body>
</html>
