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
if info(0)="" then
%>
<script language="Vbscript">
msgbox"�D�k�J�I�I",0,"ĵ�i"
history.back
</script>
<%
Response.End
end if%>
<%
'rs.open "select * from diarymessage where clicknumber>counted and user='"&info(0)&"'",conn,1,3

'if not rs.bof then
'tempcount=0

'do while not rs.eof

'tempcount=tempcount+(rs("clicknumber")-rs("counted"))
'rs.update "counted",rs("clicknumber")
'rs.movenext
'loop
'rs.close
'rs.open "select * from bardian where user='"&info(0)&"'",conn,1,3
'if not rs.bof then
'rs.update "magic",tempcount*0.01
'end if
'rs.close
'else

'rs.close
'end if
%>
<%
pagecount=9
if request.querystring("page")="" or (not isnumeric(request.querystring("page")))then
dqpage=1
else
dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if
end if

rs.open "select * from diarymessage where user='"&info(0)&"' order by id desc ",conn,1,1

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
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ү���ϡV�ڪ��d���O</title>
<link rel="StyleSheet" href="../../city/new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center"><center>
</center></div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">��-�z��e����m--�L�ȴJ��---�ڪ���O��&nbsp;</td>
<td width="10"></td>
<td class=p1 width="440">��-��e�ɶ��G<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">��-��O���g�ϡG</td>
<td width="10"></td>
<td class=p1 width="440">��-��O�\���ϡG</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td width="326">
<form method="POST" action="mydiary_post.asp">

<table border="0" width="324" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="314">
<p align="right"><select size="1" name="weather" style="font-family: Tahoma; font-size: 12px">
<option value="����" selected>����</option>
<option value="����">����</option>
<option value="�h��">�h��</option>
<option value="�B">�B</option>
<option value="��">��</option>
<option value="��">��</option>
<option value="��">��</option>
</select><select size="1" name="mood" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="10">
</td>
</tr>
<center>
<tr>
<td class=p3 width="324" colspan="2"><textarea rows="14" name="detail" cols="59" style="font-family: Tahoma; font-size: 12px"></textarea></td>
</tr>
<tr>
<td class=p3 width="324" colspan="2">�o�g��O�O�_�}���O�G<input type="radio" value="yes" name="open">�O<input type="radio" value="no" checked name="open">�_
<input type="submit" value="����" name="B2" style="font-family: Tahoma; font-size: 12px">
<input type="reset" value="���g" name="B1" style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</form>
</center>
</td>
<td width="10"></td>
<td width="440" valign="top">
<table border="0" width="436" cellspacing="1" cellpadding="0"><FORM METHOD="get" ACTION="index.asp">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="186" colspan="4">��e<%=dqpage%>���G �`<%=totalpage%>���G</td>
<td class=p2 width="164"></td>
<td class=p2 width="70"></td>
</tr></FORM>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="46">

<%if dqpage>1 then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="1">
<INPUT TYPE="submit" value="����" style="font-family: Tahoma; font-size: 12px">
</FORM>
<%else%>
<INPUT TYPE="submit" value="����" style="font-family: Tahoma; font-size: 12px">
<%end if%>

</td>
<td class=p2 width="46">
<%if dqpage>1 then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage-1%>">
<INPUT TYPE="submit" value="�W��" style="font-family: Tahoma; font-size: 12px">
</FORM>

<%else%>
<INPUT TYPE="submit" value="�W��" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<td class=p2 width="46">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage+1%>">
<INPUT TYPE="submit" value="�U��" style="font-family: Tahoma; font-size: 12px">
</FORM>


<%else%>
<INPUT TYPE="submit" value="�U��" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<td class=p2 width="48">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=totalpage%>">
<INPUT TYPE="submit" value="����" style="font-family: Tahoma; font-size: 12px">
</FORM>


<%else%>
<INPUT TYPE="submit" value="����" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<FORM METHOD=get ACTION="index.asp">

<td class=p2 width="234" colspan="2">���ġG<INPUT TYPE="text" NAME="page" size="4" maxlength=4 style="font-family: Tahoma; font-size: 12px">��<input type="submit" value="�T�w" name="B1" style="font-family: Tahoma; font-size: 12px"></td></FORM>
</tr>
</table>
<table border="0" width="436" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="190">���e���n</td>
<td class=p2 width="30">�Ѯ�</td>
<td class=p2 width="80">���</td>
<td class=p2 width="40">�߱�</td>
<td class=p2 width="50">�I���q</td>
<td class=p2 width="40">�R���H</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

rs.movenext
if not rs.eof then
nid=rs("id")
end if
rs.moveprevious
rs.moveprevious
if not rs.bof then
pid=rs("id")
end if
rs.movenext

%>
<tr>
<td class=p3 width="10"></td>
<td class=p3 width="190"><A HREF="readjump.asp?pid=<%=pid%>&nid=<%=nid%>&id=<%=rs("id")%>"><%=cutstr(rs("diarymessage"),20)%></A></td>
<td class=p3 width="30"><%=rs("diaryseason")%> </td>
<td class=p3 width="80"><%=rs("diarydate")%> </td>
<td class=p3 width="40"><%=rs("mood")%> </td>
<td class=p3 width="50"><%=rs("clicknumber")%> </td>
<td class=p2 width="40">
<FORM METHOD=POST ACTION="deldiary.asp">
<INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
<INPUT TYPE="submit" value="�R��" style="font-family: Tahoma; font-size: 12px">
</FORM>
</td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count
%>
<tr>
<td class=p3 width="10" height="23"></td>
<td class=p3 width="190" height="23"></td>
<td class=p3 width="30" height="23"> </td>
<td class=p3 width="80" height="23"> </td>
<td class=p3 width="40" height="23"> </td>
<td class=p3 width="50" height="23"> </td>
<td class=p2 width="40" height="23"> </td>
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
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">��-���U</td>
<td width="10"></td>
<td class=p1 width="440">��-��O�d�߰ϡG</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0" height="104">
<tr>
<td width="324" rowspan="5">
<table border="0" width="324" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="100%" height="19">�p���F���ܡG</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">�w��D�H�S�@�������}�F�o���ݩ�z����O���A�b�o��</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">�z�i�H�Υ��O���A��m�ͬ����C�@�ѡC�b��O���g�ϡA</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">�A�i�H��ܤѮ𱡪p�A���骺�߱��A��g�ӤH��O�F��</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">�O�}��]�m�O��ܤ�O�O�_�b��������O���߶}��C</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">�p�G�z������N���Ϋ�ĳ�G<a href="mailto:bugs@flush.com.cn">bugs@flush.com.cn</a></td>
</tr>
</table>
</td>
</tr>
<form method="get" action="datesearch.asp">
<tr>
<td width="10" align="right"></td>
<td class=p2 width="100" align="right">������d�ݡG</td>
<td class=p3 width="90"><select size="1" name="year" style="font-family: Tahoma; font-size: 12px">
<option selected value="2000">2000</option>
<option value="2001">2001</option>
<option value="2002">2002</option>
<option value="2003">2003</option>
<option value="2004">2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
</select>�~</td>
<td class=p3 width="90"><select size="1" name="startm" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="90"><select size="1" name="startd" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="70"></td>
</tr>
<tr>
<td width="10" align="right"> </td>
<td class=p2 width="100" align="right"> �ܡG</td>
<td class=p3 width="90"></td>
<td class=p3 width="90"><select size="1" name="endm" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="90"><select size="1" name="endd" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="70"></td>
</tr>
<tr>
<td width="10"></td>
<td class=p2 width="100"></td>
<td class=p3 width="90"></td>
<td class=p3 width="90"><select size="1" name="mood" style="font-family: Tahoma; font-size: 12px">
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
<td class=p3 width="90"><input type="submit" value="�d��" name="B1" style="font-family: Tahoma; font-size: 12px"></td>
<td class=p3 width="70"></td>
</tr>
</form>
<FORM METHOD="get" ACTION="keywordsearch.asp">
<tr>
<td width="10" align="right">
</td>
<td class=p2 width="100" align="right">
����r�d�ߡG</td>
<td class=p3 width="180" colspan="2"><input type="text" name="keyword" size="27" maxlength=5 style="font-family: Tahoma; font-size: 12px"> </td>
<td class=p3 width="160" colspan="2"> <input type="submit" value="�d��" name="B1" style="font-family: Tahoma; font-size: 12px"> </td>
</tr>
</form>
</table>
</center>
</div>
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
</tr></table></center></div> <%
'=====================================================

'�I���@�w���ת��r�Ŧ�A��l��" ..."
'�եΡGcutstr(�r�Ŧ�A�I������)

function cutstr(tempstr,tempwid)

if instr(tempstr,"<br>") then
tempstr=replace(tempstr,"<br>"," ")
end if

if instr(tempstr,"&nbsp;") then
tempstr=replace(tempstr,"&nbsp;"," ")
end if

if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&" ..."
else
cutstr=tempstr
end if
end function
'=====================================================

%>
</body>

</html>
<html><script language="JavaScript">                                                                  </script></html>
