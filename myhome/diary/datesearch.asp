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
<%
sdate=request.querystring("year")&"-"&request.querystring("startm")&"-"&request.querystring("startd")
edate=request.querystring("year")&"-"&request.querystring("endm")&"-"&request.querystring("endd")

if sdate>edate then
%>
<script language="Vbscript">
msgbox"����d���ܿ��~�I",0,"ĵ�i"
history.back
</script>
<%
Response.End
end if

if request.querystring("mood")="" then
sql="select * from diarymessage where user='"&info(0)&"' and diarydate>='"&sdate&"' and diarydate<='"&edate&"'"
else
sql="select * from diarymessage where user='"&info(0)&"' and diarydate>='"&sdate&"' and diarydate<='"&edate&"' and mood='"&request.querystring("mood")&"'"
end if



pagecount=9

if request.querystring("page")="" or (not isnumeric(request.querystring("page")))then
dqpage=1
else
dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if
end if

rs.open sql,conn,1,1
if rs.bof then
%>
<script language="Vbscript">
msgbox"�O�����s�b",0,"ĵ�i"
history.back
</script>
<%
end if
temp1=rs.recordcount/pagecount
if temp1-int(rs.recordcount/pagecount)<>0 then
totalpage=int(temp1+1)
else
totalpage=temp1
end if
if dqpage>totalpage then
dqpage=1
end if

if dqpage<>1 then
temp2=(dqpage-1)*pagecount
rs.move temp2

end if
if dqpage>totalpage or dqpage<0 then

%>
<script language="Vbscript">
msgbox"���X���s�b�I",0,"ĵ�i"
history.back
</script>
<%
Response.End
end if
%>
<html>

<head>
<title>�F�C����-�ڪ���O</title>
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
<font size="2" color="#E18C59"><b>�ڪ���O</b></font>
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
<div align="center" style="width: 694; height: 43">
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
<table border="0" width="686" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="354">��-�z��e����m-<font
color="#FFCC00"><a href="../../jh.asp" style="color: #FFCC00">�ӫ�����</a></font>-<a
href="../index.asp"><font color="#FFCC00">�L�ȴJ��</font></a>-<a
href="index.asp"><font color="#FFCC00">�ڪ���O</font></a>-�d�ߤ�O</td>
<td width="4"></td>
<td class="p1" width="407">��-��e�ɶ��G<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="681" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="324">��-��O���g�ϡG</td>
<td width="10"></td>
<td class="p1" width="431">��-��O�\���ϡG</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="685" cellspacing="1" cellpadding="0">
<tr>
<td width="326">
<form method="POST" action="mydiary_post222.asp">
<table border="0" width="324" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="314">
<p align="right"><select size="1" name="weather">
<option value="����" selected>����</option>
<option value="����">����</option>
<option value="�h��">�h��</option>
<option value="�B">�B</option>
<option value="��">��</option>
<option value="��">��</option>
<option value="��">��</option>
</select><select size="1" name="mood">
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
<td class="p3" width="10"></td>
</tr>
<center>
<tr>
<td class="p3" width="324" colspan="2"><textarea rows="9"
name="detail" cols="42"></textarea></td>
</tr>
<tr>
<td class="p3" width="324" colspan="2">�o�g��O�O�_�}���O�G<input
type="radio" value="yes" name="open">�O<input
type="radio" value="no" checked name="open">�_ <input
type="submit" value="����" name="B2"> <input
type="submit" value="�g���F" name="B1"></td>
</tr>
</table>
</form>
</center></td>
<td width="10"></td>
<td width="442" valign="top">
<table border="0" width="342" cellspacing="1" cellpadding="0">
<FORM METHOD="get" ACTION="datesearch.asp">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="49">���G</td>
<td class="p2" width="198" colspan="3"><input type="text"
name="page" size="19"><input type="submit" value="�T�w"
name="B1"></td>
<input type="hidden" name="startm"
value="<%=request.querystring("startm")%>"><input
type="hidden" name="startd"
value="<%=request.querystring("startd")%>"><input
type="hidden" name="endm"
value="<%=request.querystring("endm")%>"><input
type="hidden" name="endd"
value="<%=request.querystring("endd")%>"><input
type="hidden" name="year"
value="<%=request.querystring("year")%>">
<td class="p2" width="62"></td>
<td class="p2" width="43"></td>
</tr>
</form>
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="49">
<%if dqpage>1 then%>
<form method="get" action="datesearch222.asp">
<input type="hidden" name="mood"
value="<%=request.querystring("mood")%>"><input
type="hidden" name="year"
value="<%=request.querystring("year")%>"><input
type="hidden" name="startm"
value="<%=request.querystring("startm")%>"><input
type="hidden" name="startd"
value="<%=request.querystring("startd")%>"><input
type="hidden" name="endm"
value="<%=request.querystring("endm")%>"><input
type="hidden" name="endd"
value="<%=request.querystring("endd")%>"><input
type="hidden" name="page" value="1"><input type="submit"
value="����">
</form>
<%else%>
<input type="submit" value="����">
<%end if%>

</td>
<td class="p2" width="47"><%if dqpage>1 then%>
<form method="get" action="datesearch222.asp">
<input type="hidden" name="mood"
value="<%=request.querystring("mood")%>"><input
type="hidden" name="year"
value="<%=request.querystring("year")%>"><input
type="hidden" name="startm"
value="<%=request.querystring("startm")%>"><input
type="hidden" name="startd"
value="<%=request.querystring("startd")%>"><input
type="hidden" name="endm"
value="<%=request.querystring("endm")%>"><input
type="hidden" name="endd"
value="<%=request.querystring("endd")%>"><input
type="hidden" name="page" value="<%=dqpage-1%>"><input
type="submit" value="�W��">
</form>
<%else%> <input type="submit" value="�W��">
<%end if%></td>
<td class="p2" width="47"><%if dqpage<totalpage then%>
<form method="get" action="datesearch222.asp">
<input type="hidden" name="mood"
value="<%=request.querystring("mood")%>"><input
type="hidden" name="year"
value="<%=request.querystring("year")%>"><input
type="hidden" name="startm"
value="<%=request.querystring("startm")%>"><input
type="hidden" name="startd"
value="<%=request.querystring("startd")%>"><input
type="hidden" name="endm"
value="<%=request.querystring("endm")%>"><input
type="hidden" name="endd"
value="<%=request.querystring("endd")%>"><input
type="hidden" name="page" value="<%=dqpage+1%>"><input
type="submit" value="�U��">
</form>
<%else%><input type="submit" value="�U��"><%end if%></td>
<td class="p2" width="104"><%if dqpage<totalpage then%>
<form method="get" action="datesearch222.asp">
<input type="hidden" name="mood"
value="<%=request.querystring("mood")%>"><input
type="hidden" name="year"
value="<%=request.querystring("year")%>"><input
type="hidden" name="startm"
value="<%=request.querystring("startm")%>"><input
type="hidden" name="startd"
value="<%=request.querystring("startd")%>"><input
type="hidden" name="endm"
value="<%=request.querystring("endm")%>"><input
type="hidden" name="endd"
value="<%=request.querystring("endd")%>"><input
type="hidden" name="page" value="<%=totalpage%>"><input
type="submit" value="����">
</form>
<%else%><input type="submit" value="����"><%end if%></td>
<td class="p2" width="62">��e���G<%=dqpage%></td>
<td class="p2" width="43">�`���G<%=totalpage%></td>
</tr>
</table>
<table border="0" width="340" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="191">���e���n</td>
<td class="p2" width="48">�Ѯ�</td>
<td class="p2" width="59">���</td>
<td class="p2" width="48">�߱�</td>
<td class="p2" width="45">�I���q</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

%>
<tr>
<%
if len(rs("diarymessage"))>15 then
diarymessage=left(rs("diarymessage"),15)&" ..."
else
diarymessage=rs("diarymessage")
end if
%>
<td class="p3" width="10"></td>
<td class="p3" width="191"><a
href="readjump.asp?id=<%=rs("id")%>"><%=diarymessage%></a></td>
<td class="p3" width="48"><%=rs("diaryseason")%> </td>
<td class="p3" width="59"><%=rs("diarydate")%> </td>
<td class="p3" width="48"><%=rs("mood")%> </td>
<td class="p3" width="45"><%=rs("clicknumber")%> </td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count
%>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="191"></td>
<td class="p3" width="48"> </td>
<td class="p3" width="59"> </td>
<td class="p3" width="48"> </td>
<td class="p3" width="45"> </td>
</tr>
<%
next
end if
rs.close
conn.close
%>
</table>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="684" cellspacing="1" cellpadding="0"
height="104">
<tr>
<td width="324" rowspan="5">
<table border="0" width="324" cellspacing="1" cellpadding="0"
height="114">
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
<tr>
<td class="p3" width="100%" height="17">&nbsp;</td>
</tr>
</table>
</td>
</tr>
<form method="get" action="datesearch.asp">
<tr>
<td width="10" align="right"></td>
<td class="p2" width="100" align="right">������d�ݡG</td>
<td class="p3" width="90"><select size="1" name="year">
<option selected value="2000">2000</option>
<option value="2001">2001</option>
<option value="2002">2002</option>
<option value="2003">2003</option>
<option value="2004">2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
</select>�~</td>
<td class="p3" width="90"><select size="1" name="startm">
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
<td class="p3" width="90"><select size="1" name="startd">
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
<td class="p3" width="64"></td>
</tr>
<tr>
<td width="10" align="right"></td>
<td class="p2" width="100" align="right">�ܡG</td>
<td class="p3" width="90"></td>
<td class="p3" width="90"><select size="1" name="endm">
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
<td class="p3" width="90"><select size="1" name="endd">
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
<td class="p3" width="64"></td>
</tr>
<tr>
<td width="10"></td>
<td class="p2" width="100"></td>
<td class="p3" width="90"></td>
<td class="p3" width="90"><select size="1" name="mood">
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
<td class="p3" width="90"><input type="submit" value="�d��"
name="B1"></td>
<td class="p3" width="64"></td>
</tr>
</form>
<FORM METHOD="get" ACTION="keywordsearch222.asp">
<tr>
<td width="10" align="right"></td>
<td class="p2" width="100" align="right">����r�d�ߡG</td>
<td class="p3" width="180" colspan="2"><input type="text"
name="keyword" size="22"></td>
<td class="p3" width="154" colspan="2"><input type="submit"
value="�d��" name="B1"></td>
</tr>
</form>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="683" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="683">&nbsp;</td>
</tr>
</table>
</center>
</div>
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
