<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)=""  then
Response.Redirect "../error.asp?id=440"
else
randomize
m=int(6*rnd+1)
if m>3 then
if request.form("select")="big" then
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(5 * Rnd + 1)
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(7 * Rnd + 1)
if m2>6 then m3=6 end if
end if
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(6 * Rnd + 1)
end if



if request.form("select")="big" then
if m1+m2+m3>9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3>=10
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
else
if m1+m2+m3<9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3<9
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
end if


dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "�}���򪱯��H�z�U�o��֪��`��Ӥ���l�H"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "�}���򪱯��H�z��o��h�Q�s�گ}���ڡI"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="Select  �Ȩ�,�y�O from �Τ� where id="&info(9)
set rs=conn.Execute(sql)
nowcash=rs("�Ȩ�")
if betcash>nowcash then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "���S���d���H�z�Ѳ��]�̭����o��h�Ȥl�H�]�K�K�K�A�֥h���I�h�a�A�a���I�^"
location.href = "javascript:history.back()"
</script>
<%
else

nowmeili=rs("�y�O")
if nowmeili<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "�j�L�A�z�ثe���y�O�C��10�F�A���l�W�ݤ��L�h�r�A�H��A�ӧa�I"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<html>
<head>
<title>������] - ��j�p</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body text="#000000" background="../images/8.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p><font size="3" class="c" color="#000000"><b>�� �� - �� �j �p<br>
</b></font><font size="2" class="c"><b><font color="#FF0033">�A�㪺�O</font><font color="#CC0000">
<%if request.form("select")="big" then%>
</font><font color="#CC0000"> </font></b></font><img src="../jhimg/bbs/big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000">
<%else%>
<img src="../jhimg/bbs/small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000">
<%end if%>
</font></b></font>
<%if (m1+m2+m3)>9 then%>
</p>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" bordercolor="#000000">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���G:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�I�A�j</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF">
<%if request.form("select")="small" then%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>���a�G</b></font></td>
<td colspan="2"><font size="2">�����٭n�ӶܡH ���䥼����@~�I</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>�ڻ��G</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">�A�ӡI�����򤣡H�A�کA�I�ڴN���۫H�o��˾`~�I</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">��F��F�A�{�˾`�a�A�d���I�Ȥl�ϩR�a~�I</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%
conn.execute("Update �Τ� set �Ȩ�=�Ȩ�-"&betcash&",�y�O=�y�O-1 where id="&info(9))
%>
</font><font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%> </b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> ��</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>���a�G</b></font></td>
<td colspan="2"><font size="2">�F�`�r~�I�ȩx�z�٭n�ӶܡH </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>�ڻ��G</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">��M�A�ӡI�ڤ��ѹB�𤣿�~�I</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
            <td colspan="2"><font size="2"><a href="../welcome.asp">���n�N���A�A�H���ڶ̡H�K�K�I</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �Τ� set �Ȩ�=�Ȩ�+"&betcash&",�y�O=�y�O-1 where id="&info(9))
%>
</font><font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font> </td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="2"><a href="betindex.asp"><font size="2">��</font><font size="2">������</font></a></td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" height="160">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���G:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�I�A�p</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2" color="#FFFFFF">
<%if request.form("select")="big" then%>
</font> <font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>���a�G</b></font></td>
<td colspan="2"><font size="2">�����٭n�ӶܡH ���䥼����@~�I</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>�ڻ��G</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">�A�ӡI�����򤣡H�A�کA�I�ڴN���۫H�o��˾`~�I</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">��F��F�A�{�˾`�a�A�d���I�Ȥl�ϩR�a~�I</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �Τ� set �Ȩ�=�Ȩ�-"&betcash&",�y�O=�y�O-1 where id="&info(9))
%>
</font> <font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%></b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> ��</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC">
<tr>
<td align="right"><font size="2" color="#000099"><b>���a�G</b></font></td>
<td colspan="2"><font size="2">�F�`�r~�I�ȩx�z�٭n�ӶܡH </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>�ڻ��G</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">��M�A�ӡI�ڤ��ѹB�𤣿�~�I</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">���n�N���A�A�H���ڶ̡H�K�K�I</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �Τ� set �Ȩ�=�Ȩ�+"&betcash&",�y�O=�y�O-1 where id="&info(9))
%>
</font> <font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3"><a href="betindex.asp"><font size="2">�������</font></a></td>
</tr>
</table>
<%end if%>
</div>
</body>
</html>
<%rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
set username=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
end if
%>
