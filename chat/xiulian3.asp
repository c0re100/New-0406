<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>�Z�ǭ׷�</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}
</script><style>
td{font-size:9pt;}
</style>
</HEAD>
<BODY bgcolor="#ffffff" background="bk.jpg">
<TABLE border="0" cellpadding="0" cellspacing="0" width="142">
<TR> 
<TD align="center" height="10" width="20" valign="top"><IMG src="./fq/tb-a1.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a6.gif" height="10" width="102"></TD>
<TD height="10" align="center" valign="top" width="57"><IMG src="./fq/tb-a7.gif" width="20" height="20" border="0"></TD>
</TR>
<TR> 
<TD width="20" background="./fq/tb-a2.gif" height="80"></TD>
<TD width="102" height="80" align="center" valign="top" background="./fq/tb-a3.gif">
<table border="0" cellspacing="0" cellpadding="1" width="105" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
<tr bgcolor=FF6633>
<td><div align="center">�ۦ�</div></td>
<td><div align="center">����</div></td>
<td><div align="center">�׷�</div></td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �Z�\,�ŧO,�׷ү� FROM �Z�\ where ����='" & info(5) & "' ",conn
do while not rs.eof and not rs.bof
%>
<tr bgcolor="#EEFFEE"  onmouseout="this.bgColor='#EEFFEE';"onmouseover="this.bgColor='#DFEFFF';">
<td><div align="center"><%=rs("�Z�\")%> </div></td>
<td><div align="center"><%=rs("�ŧO")%> </div></td>
<td><div align="center"><a href="javascript:s('/�ײߪZ�\$ <%=rs("�Z�\")%>')" title="[<%=rs("�Z�\")%>] �ٮt<%=(20000-rs("�׷ү�"))%>�I�ɯ�!"><font color="#FF0000">�׷�</font></a></div></td>
</tr>

<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>  <tr>
<td colspan="3" height="17" bgcolor="#FF6633">
<div align="center"><font color="#FFFFFF">
<input type="checkbox" name="liji" value="on">
<font color="#000000">�ϥX������</font> </font></div>
</td>
</tr>
</table></TD>
<TD background="./fq/tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="./fq/tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a5.gif" width="102" height="10"></TD>
<TD width="57" height="10"><IMG src="./fq/tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML>