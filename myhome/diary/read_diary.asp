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
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ү���ϡV�ڪ��d���O</title>
<link rel="StyleSheet" href="../../new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center"><center>
</center></div><div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">��-�z��e����m---�L�ȴJ��---��O��---��O����-�\Ū��O</td>
<td width="10"></td>
<td class=p1 width="440">��-��e�ɶ��G<%=date%>&nbsp;&nbsp;<%=time%></td>
</tr>
</table>
</center>
</div>


<div align="center">
<center>

<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">��-��O�\���ϡG&nbsp;&nbsp; <a href="javascript:history.back(1)">��^</a></td>
<td width="10"></td>
<td class=p1 width="440">��-��O�d�߰ϡG</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">


<tr>
<td class=p2 width="10"></td>
<%rs.open "select * from diarymessage where user='"&info(0)&"' order by id desc",conn,1,1
do while not rs.eof
if cint(rs("id"))=cint(request.querystring("id")) then
rs.movenext
exit do
end if
rs.movenext
loop

%>


<td class=p2 width="40"><%
if rs.eof then
%>
<INPUT TYPE="submit" value="�U�@��" style="font-family: Tahoma; font-size: 12px">
</td>
<%else%>
<FORM METHOD=get ACTION="read_diary.asp">
<INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
<INPUT TYPE="submit" value="�U�@��" style="font-family: Tahoma; font-size: 12px"></FORM>


<%end if%>
<td class=p2 width="254"><%
rs.moveprevious
rs.moveprevious
if rs.bof then%>

<INPUT TYPE="submit" value="�W�@��" style="font-family: Tahoma; font-size: 12px">
</td>
<%else%>
<FORM METHOD=get ACTION="read_diary.asp">
<INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
<INPUT TYPE="submit" value="�W�@��" style="font-family: Tahoma; font-size: 12px"></FORM>
<%end if
rs.movenext%>

<td width="10"></td>
<td class=p2 width="440"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>

<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="70">����G</td>
<%

%>
<td class=p3 width="115"><%=rs("diarydate")%>&nbsp;&nbsp;<%=rs("diarytime")%></td>
<td class=p3 width="113">�@�̡G<%=rs("user")%></td>
<td width="10"></td>
<td width="440" rowspan="5">
<table border="0" width="440" cellspacing="1" cellpadding="0" height="111">
<form method="get" action="datesearch.asp">
<tr>
<td class=p2 width="100" align="right" height="24">������d�ݡG</td>
<td class=p3 width="90" height="24"><select size="1" name="year" style="font-family: Tahoma; font-size: 12px">
<option selected value="2000">2000</option>
<option value="2001">2001</option>
<option value="2002">2002</option>
<option value="2003">2003</option>
<option value="2004">2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
</select>�~</td>
<td class=p3 width="90" height="24"><select size="1" name="startm" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="90" height="24"><select size="1" name="startd" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="70" height="24"></td>
</tr>
<tr>
<td class=p2 width="100" align="right" height="25">�ܡG</td>
<td class=p3 width="90" height="25"></td>
<td class=p3 width="90" height="25"><select size="1" name="endm" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="90" height="25"><select size="1" name="endd" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="70" height="25"></td>
</tr>
<tr>
<td class=p2 width="100" align="right" height="27"></td>
<td class=p3 width="90" height="27"></td>
<td class=p3 width="90" height="27"><select size="1" name="mood" style="font-family: Tahoma; font-size: 12px">
<option selected value="">�Ҧ��߱�</option>
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
<td class=p3 width="90" height="27"><input type="submit" value="�d��" name="B1" style="font-family: Tahoma; font-size: 12px"></td>
<td class=p3 width="70" height="27"></td>
</tr> </form>
<tr> <FORM METHOD="get" ACTION="keywordsearch.asp">
<td class=p2 width="100" align="right" height="27">����r�d�ߡG</td>
<td class=p3 width="180" colspan="2" height="27"><input type="text" name="keyword" size="22" style="font-family: Tahoma; font-size: 12px"> </td>
<td class=p3 width="160" colspan="2" height="27"> <input type="submit" value="�d��" name="B1" style="font-family: Tahoma; font-size: 12px"> </td>
</form>
</tr>
</table>

</td>
</tr>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="70">�Ѯ�G</td>
<td class=p3 width="244" colspan="2"><%=rs("diaryseason")%></td>
<td width="10"></td>
</tr>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="70">�߱��G</td>
<td class=p3 width="244" colspan="2"><%=rs("mood")%></td>
<td width="10"></td>
</tr>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="70">�}���v���G</td>
<td class=p3 width="244" colspan="2"><%=rs("open")%></td>
<td width="10"></td>
</tr>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="70">�I���ơG</td>
<td class=p3 width="244" colspan="2"><%=rs("clicknumber")%></td>
<td width="10"></td>
</tr>
</table>


</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="312">���e�G</td>
<td width="10"></td>
<td class=p2 width="442"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="10" rowspan="2"></td>
<td class=p3 width="312" rowspan="2"><%=rs("diarymessage")%></td>
<td width="10" rowspan="2"></td>
<td class=p3 width="442" rowspan="2"></td>
</tr>

</table>
</center>
</div>
<%
rs.close
conn.close
%>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="776">&nbsp;</td>
</tr>
</table>
</center>
</div>
<hr width="776" color="#000000" size="1">
<div align="center"><center>
<table border="0" width="468" cellspacing="0" cellpadding="0">
<tr><td width="100%"><IFRAME src="../../down.htm" width="468" height="120" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="NO"></IFRAME></td>
</tr></table></center></div>
</body>

</html>