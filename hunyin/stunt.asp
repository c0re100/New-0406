<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �G�B FROM �Τ� where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("�G�B")="" or rs("�G�B")="�L" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG["&info(0)&"�A�٨S���G�B�O�A�ӧ@����I]');history.go(-1);</script>"
	response.end
end if
%>
<html>
<head>
<title>�۳Фҩd�X���</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#FFFFFF" bgcolor="#660000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"> <b>�� �d �X �� �� �] �m</b><font color="#000000"><br>
<br>
</font> </div>
<table cellpadding="0" cellspacing="1" border="1" align="center" width="597" bgcolor="#00CCFF" bordercolor="#000000">
<tr>
<td height="11" width="98">
<div align="center"> <font color="#000000" size="2">�� �� �W</font> </div>
</td>
<td height="11" width="124">
<div align="center"> <font color="#000000" size="2">�� �� �� �O</font> </div>
</td>
<td height="11" width="367">
<div align="center"> <font color="#000000" size="2">�� �@</font> </div>
</td>
</tr>
<%
rs.close
rs.open "SELECT �X���,���O FROM �X��� where �m�W�k='" & info(0) & "' or �m�W�k='" & info(0) & "'"
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
<form method=POST action='stunt1.asp?a=m&wg=<%=rs("�X���")%>'>
<tr>
<td height="2" width="98">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="wg1" size="10"
value="<%=rs("�X���")%>" maxlength="8">
</font> </div>
</td>
<td height="2" width="124">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="nl" size="10"
value="<%=rs("���O")%>" maxlength="8">
</font> </div>
</td>
<td height="2" width="367">
<div align="center"> <font color="#000000" size="2">�m�W�G
<input type="text" name="name"
size="10" maxlength="10">
�K�X�G
<input type="password" name="pass"
size="10" maxlength="20">
<input type="submit" value="�ק�"
name="submit">
</font> </div>
</td>
</tr>
</form>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<10 then
%>
<form method=POST action='stunt1.asp?a=n'>
<tr>
<td width="98" height="2">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="wg1" size="10"
maxlength="8">
</font> </div>
</td>
<td width="124" height="2">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="nl" size="10"
maxlength="8">
</font> </div>
</td>
<td width="367" height="2">
<div align="center"> <font color="#000000" size="2">�m�W�G
<input type="text" name="name"
size="10" maxlength="10">
�K�X�G
<input type="password" name="pass"
size="10" maxlength="20">
<input type="submit" value="�K�["
name="submit">
</font> </div>
</td>
</tr>
</form>
<%end if%>
</table>
<p class="p1" align="center"><font size="2">[�`�G�u���G�B�j�L�ΫL�k�ץi�H�۳СA�C�Ыؤ@������10000��A�Ыث����賣�i�ϥΡA���B��X��ޥ��ġI]<br>
  <br>
  [���ˤO=�k����ˤO+�k����ˤO]</font><font color="#000000" size="2"><br>
  <br>
  <br>
  </font></p>
</body>
</html> 