<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>������</title><meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body BGCOLOR="#996699" background="../ff.jpg" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"> 
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ��O from �Τ� where id="&info(9) ,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
  <script language=vbscript>
MsgBox "�藍�_�A����i��ާ@�I�A�٤��O���򤤤H�A�u�C"
window.close()
</script> <%
else
sm=rs("��O")
conn.close
set conn=nothing
set rs=nothing

%> <font color="#000000">�w��<%=info(0)%>�ӽ������v���]�����ƬG���t�d���^ </font> </div><table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" BACKGROUND="../IMAGES/BJ31.JPG"> <tr> <td bgcolor="#CCE8D6" BACKGROUND="../ljimage/downbg.jpg"> 
<table> <form method=POST action='yilao2.asp' onsubmit="this.ok.disabled=true"> 
<tr> <td>�J��P���G�A���ͩR�O��<%=sm%>�C <%
if sm>=1000 then response.write "���ݭn�v���I":p=1
if sm<1000 and sm>0 then response.write "��ĳ�h��Ǹɫ~�I":p=2
if sm<0 and sm>-500 then response.write "�A���馳�~�A��ĳ�v���I":p=3
if sm<-500 then response.write "�A�˱o�����A�p���v���N�M�ΥͩR�I":p=4
%> <tr> <td>�v����k�G<input name="yilao" readonly size="8"
value="<%
if p=1 then response.write "�L"
if p=2 then response.write "��Ǹɫ~"
if p=3 then response.write "�@��v��"
if p=4 then response.write "�m�ϥͩR"
%>"> <input name="ok" type="submit" value="�T�w"> <INPUT TYPE="button" VALUE="����" ONCLICK="javascript:window.close();" NAME="button"></td></tr> 
<tr> <td colspan="2" style="font-size:9pt"> <hr> �������v�����ؼг��O�N�ͩR�ը�1000�C<br> 1�B���ɥD�n�O�w��ͩR�Ȧb0--1000�����A�C��Ȥl�[�ͩR��1.5<br> 
2�B�@��v���O�w��ͩR�Ȧb-500--0�����A�C��Ȥl�[�ͩR��1.0<br> 3�B���O�m�ϬO�w��ͩR�Ȧb-500�H�U���A�C��Ȥl�[�ͩR��0.5�C</td></tr> 
</form></table></td></tr> </table><%end if%> 
</body>
</html>
