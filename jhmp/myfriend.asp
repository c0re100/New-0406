<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
%>
<html>
<head>
<title>����ԤH�O��</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body bgcolor="#006699" background="../ff.jpg" text="#000000" vlink="#FFFF00" alink="#99FFFF"
leftmargin="0" topmargin="0">
<p align="center"> <FONT COLOR="#000033">�P�§A�B�ͳo�ǤH�O�A�Ԩ�ڭ̦��򪺡I<br>
  �u���A�Ԫ��H�s�I��100�I�ץi�H�b�o����ܥX�ӡI<br>
  <br>
  ²���G�ԤH�i�H�W�[�A�ۤv���I�ơA��A�ҩԨ�ڭ̦��򪺤H�s�I�j��100��,<br>
  �A�C�Ӥ볣�i�H�q�L���W���֤@�w�I�ơ]�O�p����۰ʫ�5%�p��A�ä��v�T�ҩԤH���w�I�^�A�b�멳��<br>
  �ɭԬO�̦h���C�p�G�o�@�Ӥ�L�h�F�A�A�N������֪��A�u�����@�U�Ӥ�F�A���Ȥ��֥[�I�ҥH�Фj�a�h�ԤH�a�I</FONT> <br>
   
<table width="621" border="0" cellpadding="2" cellspacing="2" height="13" align="center" bordercolor="#990000">
  <tr bgcolor="#FF0000"> 
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,�m�W,�ʧO,����,����,lasttime,mvalue,grade,����,�O�d2 FROM �Τ� where allvalue>100 and ���ФH='"& info(0) &"' order by -mvalue",conn
%>
    <td width="26" height="10"> 
      <div align="center"><font color="#FFFFFF">ID</font></div></td>
    <td width="61" height="10"> 
      <div align="center"><font color="#FFFFFF">�m�W</font></div></td>
    <td width="26" height="10"> 
      <div align="center"><font color="#FFFFFF">�ʧO</font></div></td>
    <td width="52" height="10"> 
      <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td width="58" height="10"> 
      <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td width="113" height="10"> 
      <div align="center"><font color="#FFFFFF">�̫�n��</font></div></td>
    <td width="48" height="10"> 
      <div align="center"><font color="#FFFFFF">��w�I</font></div></td>
    <td width="42" height="10"> 
      <div align="center"><font color="#FFFFFF">�޲z��</font></div></td>
    <td width="46" height="10"> 
      <div align="center"><font color="#FFFFFF">�԰���</font></div></td>
    <td width="36" height="10"> 
      <div align="center"><font color="#FFFFFF">�����I</font></div></td>
    <td width="45" height="10"> 
      <div align="center"><font color="#FFFFFF">����</font></div></td>
  </tr>
  <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
  <tr> 
    <td width="26" height="30"> <div align="center"><font color="#000000"><%=rs("ID")%></font></div></td>
    <td width="61" height="30"> <div align="center"><font color="#0000FF"><%=rs("�m�W")%></font></div></td>
    <td width="26" height="30"> <div align="center"><font color="#000000"><%=rs("�ʧO")%></font></div></td>
    <td width="52" height="30"> <div align="center"><font color="#000000"><%=rs("����")%></font></div></td>
    <td width="58" height="30"> <div align="center"><font color="#000000"><%=rs("����")%></font></div></td>
    <td width="113" height="30"> <div align="center"><font color="#000000"><%=rs("lasttime")%></font></div></td>
    <td width="48" height="30"> <div align="center"><font color="#FF0000"><%=rs("mvalue")%></font></div></td>
    <td width="42" height="30"> <div align="center"><font color="#000000"><%=rs("grade")%></font></div></td>
    <td width="46" height="30"> <div align="center"><font color="#000000"><%=rs("����")%></font></div></td>
    <td width="36" height="30"> <div align="center"><font color="#FF0000"><%=int(rs("mvalue")*0.05)%></font></div></td>
    <td width="45" height="30"> <div align="center"><font color="#000000"> 
        <%
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,sj)=0 and rs("�O�d2")<>"����"&Month(date()) and rs("mvalue")>20 then%>
        <a href="babi.asp?id=<%=rs("id")%>"><font color="#0000FF">����</font></a> 
        <%else%>
        ����ާ@ 
        <%end if%>
        </font></div></td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p align="center"> <font color="#000000"><br> <font size="+3">�ԤH�`��:</font></font><font size="+3"><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">�H</font></font> 
<div align="center"></div>
</body>
</html>