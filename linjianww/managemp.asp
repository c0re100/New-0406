<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
id=Request.QueryString("ID")
sql="Select * from ���� where ����='"&Id&"'"
rs=conn.execute(sql)
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FAF0E2" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="0064.jpg">
<div align="center"> <b><font color="#FF0000" size="+6">�������e�ק�</font></b></div>
<form action="updatmp.asp?subject=<%=rs("����")%>" method=POST>
  <ul>
    <table border=1 cellspacing=0 cellpadding=2 align="center" width="533" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
      <tr> 
        <td width="99"><font size="2" class="c" color="#FF6600">����</font></td>
        <td width="420"> 
          <input name="subject" size=40 maxlength=50 value="<%=RS("����")%>">
          <font color="#FF6600">�����W���o�j��10�ӡI </font></td>
      </tr>
      <tr> 
        <td width="99"><font color="#FF6600">²��</font></td>
        <td width="420"> 
          <input name="comment" value="<%=RS("²��")%>" size=40 maxlength=50>
          <font color="#FF6600">�r�Ƥ��W�L50�ӡI</font></td>
      </tr>
      <tr> 
        <td width="99"><font class="c" size="2" color="#FF6600">�x��&nbsp;</font></td>
        <td width="420"> 
          <input name="name" value="<%=RS("�x��")%>" size=40 maxlength=50>
          <font color="#FF6600">���]�x���ж�g���w�I</font> </td>
      </tr>
      <tr> 
        <td width="99"><font class="c" size="2" color="#FF6600">�������</font></td>
        <td width="420"> 
          <input name="jijing" value="<%=RS("�������")%>" size=40 maxlength=50>
          <font color="#FF6600"> ���o�W20���I</font> </td>
      </tr>
      <tr> 
        <td width="99"><font color="#FF6600">�A�X</font></td>
        <td width="420"> 
          <input name="shihe" value="<%=RS("�A�X")%>" size=40 maxlength=100>
          <font color="#FF6600">(�k/�k/�Ҧ�)</font> </td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c" color="#000000">�@�ӤH�����@�Ӫ������x���A���ର�h�a���ڡA�H�̷s���]�m���ǡC<br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="��s">
        </font></p>
    </div>
  </ul>
</form>
</body>
</html>
<%
set rs=nothing
%> 