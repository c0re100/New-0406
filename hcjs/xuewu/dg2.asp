<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
my=info(0)
rs.open "select �ާ@�ɶ�,��O,���O from �Τ� where id=" & info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�t�έ���A�е�["&s&"����]�A�ާ@�I');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if	
if rs.eof or rs.bof then
	response.write "�A���O���򤤤H�A����i�J����Z�]"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else
tl=rs("��O")
nl=rs("���O")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
if tl<300 or nl<350 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "���~�I�A�ثe����O�Τ��O�Ȥ����A�^�h�h�[�m�m���B���˦A�ӧa�I"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update �Τ� set ��O=��O-300,���O=���O-350,�Z�\=�Z�\+60,���s=���s+2,�ާ@�ɶ�=now() where id="&info(9)
	rs.close
rs.open "select ����,�Z�\,�Z�\�[,���s,���s�[ from �Τ� where id=" & info(9),conn
wgj=(rs("����")*1280+3800)+rs("�Z�\�[")
fyj=(rs("����")*1120+3000)+rs("���s�[")
if rs("�Z�\")>wgj then
conn.execute "update �Τ� set �Z�\=" & wgj & ",�ާ@�ɶ�=now() where id="&info(9)
end if
if rs("���s")>fyj then
conn.execute "update �Τ� set ���s=" & fyj & ",�ާ@�ɶ�=now() where id="&info(9)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
message="" & my & "�����W�W�ǤF�@�ѡA�A��o�Z�\60���s2�I������O300���O350�I"
		
end if
end if%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>����Z�]:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� �^" onclick="location.href='xuetang.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
