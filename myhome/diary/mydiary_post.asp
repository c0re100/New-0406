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

<!--#include file="../data2.asp"-->
<html>

<head>
<title>�F�C����-�d���ө�</title>
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
</head>

<body bgcolor="#000000" text="#FFFFFF">

<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
<td width="744" background="../../jh/tiao20.gif" height="83">
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
<font size="2" color="#E18C59"><b>�\Ū��O </b></font>
</div>
</td>
<td width="104">
<div align="center">

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
<td width="17" background="../../jh/tiao10.gif">&nbsp;</td>
<td valign="top" width="639">
<div align="center" style="width: 686; height: 43">
<div align="center">
<center>
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">
</center>
</table>
</div>
</div>
<div align="center">
<center>
<table border="1" width="695" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="324">��-�z��e����m-<a
href="index.asp"><font color="#FFCC00">�L�ȴJ��</font></a>-��O��-�\Ū��O</td>
<td width="10"> </td>
<td class="p1" width="440">��-��e�ɶ��G<%=date%>&nbsp;&nbsp;<%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="324">��-��O�\���ϡG&nbsp;&nbsp; <a
href="javascript:history.back(1)"><font color="#FFCC00">��^</font></a></td>
<td width="10"> </td>
<td class="p1" width="440">��-��O�d�߰ϡG</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="10"> </td>
<%rs.open "select * from diarymessage where user='"&info(0)&"' order by id desc",conn,1,1
do while not rs.eof
if cint(rs("id"))=cint(request.querystring("id")) then
rs.movenext
exit do
end if
rs.movenext
loop

%>


<td class="p2" width="40"><%
if rs.eof then
%>
<input type="submit" value="�U�@��"
style="font-family: Tahoma; font-size: 12px"></td>
<%else%>
<FORM METHOD=get ACTION="read_diary.asp">
<input type="hidden" name="id" value="<%=rs("id")%>">
<td><input type="submit" value="�U�@��"
style="font-family: Tahoma; font-size: 12px"></FORM>


<%end if%>
<td class="p2" width="254"><%
rs.moveprevious
rs.moveprevious
if rs.bof then%>

<input type="submit" value="�W�@��"
style="font-family: Tahoma; font-size: 12px"></td>
<%else%>
<FORM METHOD=get ACTION="read_diary.asp">
<input type="hidden" name="id" value="<%=rs("id")%>">
<td><input type="submit" value="�W�@��"
style="font-family: Tahoma; font-size: 12px"></FORM>
<%end if
rs.movenext%>

<td width="10"> </td>
<td class="p2" width="440"> </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="91">����G</td>
<%

%>
<td class="p3" width="94"><%=rs("diarydate")%>&nbsp;&nbsp;<%=rs("diarytime")%></td>
<td class="p3" width="113">�@�̡G<%=rs("user")%></td>
<td width="10"> </td>
<td width="440" rowspan="5">
<table border="1" width="440" cellspacing="1" cellpadding="0"
height="111" bordercolor="#E18C59">
<form method="get" action="datesearch.asp">
<tr>
<td class="p2" width="100" align="right" height="24">������d�ݡG</td>
<td class="p3" width="90" height="24"><select size="1"
name="year" style="font-family: Tahoma; font-size: 12px">
<option selected value="2000">2000</option>
<option value="2001">2001</option>
<option value="2002">2002</option>
<option value="2003">2003</option>
<option value="2004">2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
</select>�~</td>
<td class="p3" width="90" height="24"><select size="1"
name="startm" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>��</td>
<td class="p3" width="90" height="24"><select size="1"
name="startd" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="25">25</option>
<option value="26">26</option>
<option value="27">27</option>
<option value="28">28</option>
<option value="29">29</option>
<option value="30">30</option>
<option value="31">31</option>
</select>��</td>
<td class="p3" width="70" height="24"></td>
</tr>
<tr>
<td class="p2" width="100" align="right" height="25">�ܡG</td>
<td class="p3" width="90" height="25"></td>
<td class="p3" width="90" height="25"><select size="1"
name="endm" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>��</td>
<td class="p3" width="90" height="25"><select size="1"
name="endd" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="25">25</option>
<option value="26">26</option>
<option value="27">27</option>
<option value="28">28</option>
<option value="29">29</option>
<option value="30">30</option>
<option value="31">31</option>
</select>��</td>
<td class="p3" width="70" height="25"></td>
</tr>
<tr>
<td class="p2" width="100" align="right" height="27"></td>
<td class="p3" width="90" height="27"></td>
<td class="p3" width="90" height="27"><select size="1"
name="mood" style="font-family: Tahoma; font-size: 12px">
<option selected value>�Ҧ��߱�</option>
<option value="���q">���q</option>
<option value="���e">���e</option>
<option value="�ּ�">�ּ�</option>
<option value="����">����</option>
<option value="���">���</option>
<option value="�d��">�d��</option>
<option value="����">����</option>
<option value="�L��">�L��</option>
<option value="�h�W">�h�W</option>
<option value="�I��">�I��</option>
<option value="�P��">�P��</option>
<option value="����">����</option>
<option value="�ץ�����">�ץ�����</option>
<option value="����">����</option>
<option value="��">��</option>
<option value="�L�U">�L�U</option>
<option value="�ˤ�">�ˤ�</option>
</select></td>
<td class="p3" width="90" height="27"><input type="submit"
value="�d��" name="B1"
style="font-family: Tahoma; font-size: 12px"></td>
<td class="p3" width="70" height="27"></td>
</tr>
</form>
<tr>
<FORM METHOD="get" ACTION="keywordsearch.asp">
<td class="p2" width="100" align="right" height="27">����r�d�ߡG</td>
<td class="p3" width="180" colspan="2" height="27"><input
type="text" name="keyword" size="22"
style="font-family: Tahoma; font-size: 12px"></td>
<td class="p3" width="160" colspan="2" height="27"><input
type="submit" value="�d��" name="B1"
style="font-family: Tahoma; font-size: 12px"></td>
</form>
</tr>
</table>
</td>
</tr>
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="91">�Ѯ�G</td>
<td class="p3" width="223" colspan="2"><%=rs("diaryseason")%></td>
<td width="10"> </td>
</tr>
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="91">�߱��G</td>
<td class="p3" width="223" colspan="2"><%=rs("mood")%></td>
<td width="10"> </td>
</tr>
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="91">�}���v���G</td>
<td class="p3" width="223" colspan="2"><%=rs("open")%></td>
<td width="10"> </td>
</tr>
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="91">�I���ơG</td>
<td class="p3" width="223" colspan="2"><%=rs("clicknumber")%></td>
<td width="10"> </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="312">���e�G</td>
<td width="10"> </td>
<td class="p2" width="442"> </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p3" width="10" rowspan="2"> </td>
<td class="p3" width="312" rowspan="2"><%=rs("diarymessage")%></td>
<td width="10" rowspan="2"> </td>
<td class="p3" width="442" rowspan="2"> </td>
</tr>
</table>
</center>
</div>
<%
rs.close
conn.close
%>
</div>
</td>
<td width="25" background="../../jh/tiao10.gif"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<div align="right">
<img src="../../jh/tiao19.gif" width="728" height="31">
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>