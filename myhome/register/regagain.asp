<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

<!--#include file="data2.asp"-->
<html>

<head>
<title>�F�C����-�`�U�ԲӸ��</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script language="JavaScript">
<!--
function view(newsfile){var gt=unescape('%3e');var popup=null;var over="Launch Pop-up Navigator";popup=window.open('','popupnav','width=510,height=470,resizable=1,scrollbars=yes');if(popup!=null){if(popup.opener==null){popup.opener=self;}popup.location.href=newsfile;}}//-->
<!--
function GotoURL(value)
{
url=value
location.href=url
}
// -->
</script>

</head>

<body bgcolor="#0099FF" text="#FFFF00" background="../../jhimg/BK_HC3W.GIF">
<%
rs.open"select * from userinfo where user='"&info(0)&"'",conn,1,1

skin=rs("skin")
if skin<>"" then
skin=rs("skin")
else
skin="01.gif"
end if

rs.close
rs.open"select * from userinfo where user='"&info(0)&"'",conn,1,1
%>
<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
    <td width="744" height="83"> 
      <table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"><span class="zilong"><font color="#FFCC00">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %>
</font></span></font></font></font></td>
<td width="475" height="26">
<div align="center">
<font size="2" color="#E18C59"><b>��g�ԲӸ��</b></font>
</div>
</td>
<td width="104">
<div align="center">
<a href="../../jh.asp" target="_top"><font color="#E18C59">��^�ӫ�����</font></a>
</div>


</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
    </td>
</tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
    <td width="17">&nbsp;</td>
<td valign="top" width="639">
<div align="center">
<div align="center">
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">

</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="695" cellspacing="1" cellpadding="0">
<tr>
<td width="690">
<form method="POST" action="reg.asp">
<div align="center">
<center>
<table border="1" width="690" cellspacing="1"
cellpadding="0" bordercolor="#E18C59">
<tr>
<td width="496" class="p2" colspan="4">�p�G�z�Q�����ӫ��������~���]�������Ҫ��G�^�A�бz�{�u��g�H�U�H���G</td>
<td width="13" rowspan="23"> </td>
</tr>
<tr>
<td width="69" class="p2">�u��m�W�G</td>
<td width="196" class="p2"><input type="text"
name="name" size="28"
style="font-family: Tahoma; font-size: 12px"
value="<%=rs("name")%>" maxlength="20"> <font
color="#FF0000">*</font></td>
<td width="71" class="p2">�u��ʧO�G</td>
<td width="156" class="p2"><select name="sex" size="1"
style="font-family: Tahoma; font-size: 12px">
<option value="�k"
<%
if instr(rs("sex"),"�k") then
response.write "selected"
end if
%>>�k</option>
<option value="�k"
<%
if instr(rs("sex"),"�k") then
response.write "selected"
end if
%>>�k</option>
</select></td>
</tr>
<tr>
<td width="69" class="p3">�X�ͤ���G</td>
<td width="423" class="p3" colspan="3"><select size="1"
name="birthyear"
style="font-family: Tahoma; font-size: 12px">
<%for i = 1900 to 2000%>
<option value="<%=i%>"
<%
if int(left(rs("birthday"),4))=i then
response.write "selected"
end if
%>><%=i%></option>
<%next%>
</select> �~ <select size="1" name="birthmonth"
style="font-family: Tahoma; font-size: 12px">
<option value="01" <%
if mid(rs("birthday"),6,2)="01" then
response.write "selected"
end if
%>>01</option>
<option value="02"
<%
if mid(rs("birthday"),6,2)="02" then
response.write "selected"
end if
%>>02</option>
<option value="03"
<%
if mid(rs("birthday"),6,2)="03" then
response.write "selected"
end if
%>>03</option>
<option value="04"
<%
if mid(rs("birthday"),6,2)="04" then
response.write "selected"
end if
%>>04</option>
<option value="05"
<%
if mid(rs("birthday"),6,2)="06" then
response.write "selected"
end if
%>>05</option>
<option value="06"
<%
if mid(rs("birthday"),6,2)="06" then
response.write "selected"
end if
%>>06</option>
<option value="07"
<%
if mid(rs("birthday"),6,2)="07" then
response.write "selected"
end if
%>>07</option>
<option value="08"
<%
if mid(rs("birthday"),6,2)="08" then
response.write "selected"
end if
%>>08</option>
<option value="09"
<%
if mid(rs("birthday"),6,2)="09" then
response.write "selected"
end if
%>>09</option>
<option value="10"
<%
if mid(rs("birthday"),6,2)="10" then
response.write "selected"
end if
%>>10</option>
<option value="11"
<%
if mid(rs("birthday"),6,2)="11" then
response.write "selected"
end if
%>>11</option>
<option value="12"
<%
if mid(rs("birthday"),6,2)="12" then
response.write "selected"
end if
%>>12</option>
</select> �� <select size="1" name="birthdate"
style="font-family: Tahoma; font-size: 12px">
<%for j = 1 to 31%>
<option
value="<%
if j<10 then
putday="0"&trim(cstr(j))
else
putday=trim(cstr(j))
end if
response.write putday
%>"
<%
if int(mid(rs("birthday"),9,2))=j then
response.write "selected"
end if
%>><%=putday%></option>
<%next%>
</select> �� <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">�q�l�l�c�G</td>
<td width="200" class="p2"><input type="text"
name="email" size="35"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("email")%>"> <font
color="#FF0000">*</font></td>
<td width="73" class="p2" rowspan="2" align="center"><img
src="skin/<%=skin%>" alt="<%=session("user")%>"
width="40" height="40"></td>
<td width="158" class="p2"><a href="photo.asp"
target="_blank"><font color="#FFCC00">���ڪ��Y��</font></a></td>
</tr>
<tr>
<td width="69" class="p3">�ӤH�D���G</td>
<td width="200" class="p3"><input type="text"
name="homepage" size="35"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("homepage")%>"> <font
color="#FF0000">*</font></td>
<td width="158" class="p3"> </td>
</tr>
<tr>
<td width="69" class="p2">�a�x�a�}�G</td>
<td width="423" class="p2" colspan="3"><input
type="text" name="address" size="78"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("address")%>"></td>
</tr>
<tr>
<td width="69" class="p3">�l�F�s�X�G</td>
<td width="194" class="p3"><input type="text" name="pc"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("pc")%>"> <font
color="#FF0000">*</font></td>
<td width="68" class="p3">�pô�q�ܡG</td>
<td width="153" class="p3"><input type="text" name="tel"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="25" value="<%=rs("tel")%>"> <font
color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">¾�~�G</td>
<td width="194" class="p2"><input type="text" name="job"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="25" value="<%=rs("job")%>"></td>
<td width="68" class="p2">�����Ҹ��G</td>
<td width="153" class="p2"><input type="text"
name="idcard" size="23"
style="font-family: Tahoma; font-size: 12px"
maxlength="20" value="<%=rs("idcard")%>"> <font
color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">�u�@���G</td>
<td width="423" class="p3" colspan="3"><input
type="text" name="corporation" size="78"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("corporation")%>"></td>
</tr>
<tr>
<td width="69" class="p2">�리�J�G</td>
<td width="194" class="p2"><select size="1"
name="salary"
style="font-family: Tahoma; font-size: 12px">
<option value="�L���J" selected>�L���J</option>
<option value="500���H�U">500���H�U</option>
<option value="500�V1000��">500�V1000��</option>
<option value="1000�V2000��">1000�V2000��</option>
<option value="2000�V4000��">2000�V4000��</option>
<option value="4000�V8000��">4000�V8000��</option>
<option value="8000���H�W">8000���H�W</option>
</select></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p3">�W���覡�G</td>
<td width="423" class="p3" colspan="3">
<table border="0" width="100%" cellspacing="1"
cellpadding="0">
<tr>
<td width="6%"><input type="checkbox" name="mode1"
value="�a��"
<%
if instr(rs("mode"),"�a��") then
response.write "checked"
end if
%>></td>
<td width="26%">�a��</td>
<td width="6%"><input type="checkbox" name="mode2"
value="���"
<%
if instr(rs("mode"),"���") then
response.write "checked"
end if
%>></td>
<td width="28%">���</td>
<td width="6%"><input type="checkbox" name="mode3"
value="���a"
<%
if instr(rs("mode"),"���a") then
response.write "checked"
end if
%>></td>
<td width="28%">���a</td>
</tr>
</table>
</td>
</tr>
<tr>
<td width="69" class="p2">�����G</td>
<td width="194" class="p2"><input type="text"
name="stature" size="11"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("stature")%>"> ��� <font
color="#FF0000">*</font></td>
<td width="68" class="p2">�B�ê��p�G</td>
<td width="153" class="p2"><select size="1"
name="marriage"
style="font-family: Tahoma; font-size: 12px">
<option value="���B"
<%
if instr(rs("marriage"),"���B") then
response.write "selected"
end if
%>>���B</option>
<option value="�w�B"
<%
if instr(rs("marriage"),"�w�B") then
response.write "selected"
end if
%>>�w�B</option>
<option value="���B"
<%
if instr(rs("marriage"),"���B") then
response.write "selected"
end if
%>>���B</option>
<option value="�స"
<%
if instr(rs("marriage"),"�స") then
response.write "selected"
end if
%>>�స</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">�魫�G</td>
<td width="194" class="p3"><input type="text"
name="weight" size="7"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("weight")%>"> ���� <font
color="#FF0000">*</font></td>
<td width="68" class="p3">�Ҧb�����G</td>
<td width="153" class="p3"><select size="1" name="city"
style="font-family: Tahoma; font-size: 12px">
<option value="�|�t"
<%
if instr(rs("city"),"�|�t") then
response.write "selected"
end if
%>>�|�t</option>
<option value="�Ѭz"
<%
if instr(rs("city"),"�Ѭz") then
response.write "selected"
end if
%>>�Ѭz</option>
<option value="�_��"
<%
if instr(rs("city"),"�_��") then
response.write "selected"
end if
%>>�_��</option>
<option value="�s�F"
<%
if instr(rs("city"),"�s�F") then
response.write "selected"
end if
%>>�s�F</option>
<option value="�W��"
<%
if instr(rs("city"),"�W��") then
response.write "selected"
end if
%>>�W��</option>
<option value="�s��"
<%
if instr(rs("city"),"�s��") then
response.write "selected"
end if
%>>�s��</option>
<option value="���n"
<%
if instr(rs("city"),"���n") then
response.write "selected"
end if
%>>���n</option>
<option value="��n"
<%
if instr(rs("city"),"��n") then
response.write "selected"
end if
%>>��n</option>
<option value="��_"
<%
if instr(rs("city"),"��_") then
response.write "selected"
end if
%>>��_</option>
<option value="����"
<%
if instr(rs("city"),"����") then
response.write "selected"
end if
%>>����</option>
<option value="��Ĭ"
<%
if instr(rs("city"),"��Ĭ") then
response.write "selected"
end if
%>>��Ĭ</option>
<option value="����"
<%
if instr(rs("city"),"����") then
response.write "selected"
end if
%>>����</option>
<option value="�֫�"
<%
if instr(rs("city"),"�֫�") then
response.write "selected"
end if
%>>�֫�</option>
<option value="�w��"
<%
if instr(rs("city"),"�w��") then
response.write "selected"
end if
%>>�w��</option>
<option value="���n"
<%
if instr(rs("city"),"���n") then
response.write "selected"
end if
%>>���n</option>
<option value="�Q�{"
<%
if instr(rs("city"),"�Q�{") then
response.write "selected"
end if
%>>�Q�{</option>
<option value="���y"
<%
if instr(rs("city"),"���y") then
response.write "selected"
end if
%>>���y</option>
<option value="�e�n"
<%
if instr(rs("city"),"�e�n") then
response.write "selected"
end if
%>>�e�n</option>
<option value="�e�_"
<%
if instr(rs("city"),"�e�_") then
response.write "selected"
end if
%>>�e�_</option>
<option value="�s�F"
<%
if instr(rs("city"),"�s�F") then
response.write "selected"
end if
%>>�s�F</option>
<option value="�s��"
<%
if instr(rs("city"),"�s��") then
response.write "selected"
end if
%>>�s��</option>
<option value="���"
<%
if instr(rs("city"),"���") then
response.write "selected"
end if
%>>���</option>
<option value="�N�L"
<%
if instr(rs("city"),"�N�L") then
response.write "selected"
end if
%>>�N�L</option>
<option value="���s��"
<%
if instr(rs("city"),"���s��") then
response.write "selected"
end if
%>>���s��</option>
<option value="����"
<%
if instr(rs("city"),"����") then
response.write "selected"
end if
%>>����</option>
<option value="�C��"
<%
if instr(rs("city"),"�C��") then
response.write "selected"
end if
%>>�C��</option>
<option value="����"
<%
if instr(rs("city"),"����") then
response.write "selected"
end if
%>>����</option>
<option value="�sæ"
<%
if instr(rs("city"),"�sæ") then
response.write "selected"
end if
%>>�sæ</option>
<option value="�̵�"
<%
if instr(rs("city"),"�̵�") then
response.write "selected"
end if
%>>�̵�</option>
<option value="��L"
<%
if instr(rs("city"),"��L") then
response.write "selected"
end if
%>>��L</option>
<option value="���X�j"
<%
if instr(rs("city"),"���X�j") then
response.write "selected"
end if
%>>���X�j</option>
<option value="�O�W"
<%
if instr(rs("city"),"�O�W") then
response.write "selected"
end if
%>>�O�W</option>
<option value="9��"
<%
if instr(rs("city"),"9��") then
response.write "selected"
end if
%>>9��</option>
<option value="�D��"
<%
if instr(rs("city"),"�D��") then
response.write "selected"
end if
%>>�D��</option>
<option value="�䥦"
<%
if instr(rs("city"),"�䥦") then
response.write "selected"
end if
%>>�䥦</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">���ڡG</td>
<td width="194" class="p2"><input type="text"
name="nation" size="7"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("nation")%>"> �� <font
color="#FF0000">*</font></td>
<td width="68" class="p2">�Ǿ��G</td>
<td width="153" class="p2"><select name="education"
size="1" style="font-family: Tahoma; font-size: 12px">
<option value="�j��"
<%
if instr(rs("education"),"�j��") then
response.write "selected"
end if
%>>�j��</option>
<option value="�p��"
<%
if instr(rs("education"),"�p��") then
response.write "selected"
end if
%>>�p��</option>
<option value="�줤"
<%
if instr(rs("education"),"�줤") then
response.write "selected"
end if
%>>�줤</option>
<option value="����"
<%
if instr(rs("education"),"����") then
response.write "selected"
end if
%>>����</option>
<option value="�Ӥh"
<%
if instr(rs("education"),"�Ӥh") then
response.write "selected"
end if
%>>�Ӥh</option>
<option value="�դh"
<%
if instr(rs("education"),"�դh") then
response.write "selected"
end if
%>>�դh</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">ICQ���X�G</td>
<td width="194" class="p3"><input type="text" name="icq"
size="19" style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("icq")%>"></td>
<td width="68" class="p3">���d���p�G</td>
<td width="153" class="p3"><select size="1"
name="health"
style="font-family: Tahoma; font-size: 12px">
<option value="���d"
<%
if instr(rs("health"),"���d") then
response.write "selected"
end if
%>>���d</option>
<option value="�}�n"
<%
if instr(rs("health"),"�}�n") then
response.write "selected"
end if
%>>�}�n</option>
<option value="�@��"
<%
if instr(rs("health"),"�@��") then
response.write "selected"
end if
%>>�@��</option>
<option value="�t"
<%
if instr(rs("health"),"�t") then
response.write "selected"
end if
%>>�t</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">OICQ���X�G</td>
<td width="194" class="p2"><input type="text"
name="oicq" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("oicq")%>"></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p3">�M�I�G</td>
<td width="194" class="p3"><input type="text"
name="pager" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("pager")%>"></td>
<td width="68" class="p3"> </td>
<td width="153" class="p3"> </td>
</tr>
<tr>
<td width="69" class="p2">���ʹq�ܡG</td>
<td width="194" class="p2"><input type="text"
name="mobilephone" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("mobilephone")%>"></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p2" valign="top">�ӤH�R�n�G</td>
<td width="415" class="p2" colspan="3"><textarea
rows="3" name="preference" cols="81"
style="font-family: Tahoma; font-size: 12px"><%=rs("preference")%></textarea>
</tr>
<tr>
<td width="69" class="p2" valign="top">�ӤH���СG<br>
<font color="#FF0000">*</font></td>
<td width="415" class="p2" colspan="3"><textarea
rows="7" name="introduction" cols="81"
style="font-family: Tahoma; font-size: 12px"><%=rs("introduction")%></textarea></td>
</tr>
<tr>
<td width="69" class="p3" valign="top">�O�_�}��z���ӤH�H��</td>
<td width="415" class="p3" colspan="3"><input
type="radio" value="�O" name="open"
<%
if instr(rs("open"),"�O") then
response.write "checked"
end if
%>> �O <input type="radio" value="�_" name="open"
<%
if instr(rs("open"),"�_") then
response.write "checked"
end if
%>> �_</td>
</tr>
<tr>
<td width="69" class="p2" valign="top"> </td>
<td width="415" class="p2" colspan="3"><input
type="submit" value="�ק�" name="B1"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>
</center>
</div>
</div>
</td>
    <td width="25"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
      <div align="right"> </div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>
<%
rs.close
conn.close
%>