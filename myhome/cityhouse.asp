<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
%><!--#include file="data.asp"--><%
'rs.open "select * from �Τ� where id="&info(9),conn
set rs1=conn1.execute("select house,bigarea from userinfo where user='"&info(0)&"'")
If Rs1.Bof OR Rs1.Eof Then
hh="0"
conn1.execute"insert into userinfo(user,house) values('"&info(0)&"','"&hh&"')"
end if
rs1.close
%>
<html>

<head>
<title>�F�C����-�ЫΥ����</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="7" bgcolor="#FFFFFF">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../images/fw.gif" width="101" height="304"></td>
<td colspan="2" valign="top" align="center">
<table width="61%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="90" height="26">&nbsp;<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %> </td>
<td width="302" height="26" align="center"><font color="#FF6600">�� ��
�� �� ��</font></td>
<td width="73">
<div align="center"> <a href="javascript:history.back()">��^</a></div>
</td>
</tr>
</table>
<br>
<form method="POST" action="housedispose.asp">
<div align="center">
<center>
<table border="0" cellspacing="3" cellpadding="0" width="650">
<tr>
<td class="p3" colspan="3">
<input
type="radio" name="S1" value="buy" checked>
��- �R
<input
type="radio" name="S1" value="sale">
��- ��
<select size="1"
name="area">
<option value="�L�s">�L�s��</option>
<option value="����">���ɰ�</option>
<option value="�ڿ�">�ڿ���</option>
<option value="�Ѫ�">�Ѫư�</option>
<option value="�]�t">�]�t��</option>
<option value="���m">���m��</option>
</select>
</td>
</tr>
<tr>
<td class="p1" colspan="3"><font color="#0000FF">��-�Юھڱz������R�n�M�ͬ��ؼп�ܵ�ϡI</font></td>
</tr>
<tr>
<td class="p3" colspan="2">��-�L�s��</td>
<td class="p3" width="552">�خL���s,�z�O�ڭ̵خL�q�j�줵���^���H���C</td>
</tr>
<tr>
<td class="p2" colspan="2">��-���ɰ�</td>
<td class="p2" width="552">�ⶳ����,�z�O�O�H�Ҥ񤣤W���_��,����,���סC</td>
</tr>
<tr>
<td class="p3" colspan="2">��-�ڿ���</td>
<td class="p3" width="552">�ڸ̵̨},�z���w�۷Q,�z�O���ܯ��ܤ����H�C</td>
</tr>
<tr>
<td class="p2" colspan="2">��-�Ѫư�</td>
<td class="p2" width="552">�Ѱ����,�z���w���,�C���O�a��Ȫe���L���Q�̡C</td>
</tr>
<tr>
<td class="p3" colspan="2">��-�]�t��</td>
<td class="p3" width="552">�]&quot;�t&quot;�u�u,�z�O�ʸ̬D�@�����,&quot;�l��&quot;(?)�C</td>
</tr>
<tr>
<td class="p2" colspan="2">��-���m��</td>
<td class="p2" width="552">&quot;��&quot;�ҩ^�m,�z�O�q�q�L�D���H,����������֤F�A�̡C</td>
</tr>
<tr>
<td class="p1" colspan="3"><font color="#FF0000">��-�Юھڱz���g�ٱ��p��ܤU�����ЫΡG�^--�ڭ̥i�����b���@
�I</font></td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house1" name="R1" checked>
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h01.gif"></td>
<td class="p2" valign="top" width="552">��- [²���T��] ��Ī���B�_���s�\�γ���²���Фl�A�O�h���]���̨嫬���Фl�C<br>
����-[<font color="#FF0000">800</font>��]</td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house2" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h02.gif"></td>
<td class="p3" valign="top" width="552">��- [�@�륭��] �u���@�h���p�˩СA�ܦh��q�h���]�h�X�Ӫ��H��b�o�̡C<br>
����-[<font color="#FF0000">20000</font>��]</td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house3" name="R1">
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h03.gif"></td>
<td class="p2" width="552">��- [���q���J] �h�֬O���ä[�s�b�]�p�ӫئ����ؿv��,�e�Τg�a�Ŷ�,�q�`���γ�,�h�b��������]���,�@����v�B�ܮw�B�u�t�B���b��שΨ�L���Ϊ��ؿv���C����-[<font
color="#FF0000">50000</font>��]</td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house4" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h04.gif"></td>
<td class="p3" valign="top" width="552">��- [���ؤ��J] �ӫ������h���J�̳���b�o�̡A�j�h�O�U�������I¾�쪺�ιB�����n�����I�̡C&nbsp;
����-[<font color="#FF0000">200000</font>��]</td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house5" name="R1">
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h05.gif"></td>
<td class="p2" valign="top" width="552">��- [���q�O��] �צb�����a�q���O�֡A��q��K�A�����u���C��b�o�̪����j�h�O�U���x���ΰө����ѪO�@�I����-[<font
color="#FF0000">500000</font>��] </td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house6" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h06.gif"></td>
<td class="p3" valign="top" width="552">��- [���اO��]-�b�����ϩΦb���ϫسy���ѥ�i�����اO�֡A�ӫ������ѡB�x���M�j�I�έ̦�b�o�̡C&nbsp;&nbsp;
����-[<font color="#FF0000">1000000</font>��] </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="468" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="462" height="13">
<p align="center">
<input type="submit" value="�M�w"
name="B1" style="font-size: 9pt">
<input type="reset"
value="�Ҽ{�@�U" name="B2" style="font-size: 9pt">
</td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>

</body>

</html>
