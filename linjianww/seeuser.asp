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
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<title>�Τ�ƾڮw�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#FFFFFF" topmargin="0" background="0064.jpg">
<p align="center"><font size="+6" color="#FF0000">�F�C����ƾڮw</font></p>
<p align="left">�п�J����d�ߥΤ�G<br>
  �p�G����='�x��' �d�ߨ������x�����Τ�C<br>
  �A�p�G �Ȩ�&gt;10000 and �Z�\&gt;10000 �d�ݻȨ�j��10000�åB�Z�\�j��10000���Τ�C<br>
  �b�d�ߤ��i�H�ϥΡG&quot;and&quot; &quot;or&quot; &quot;&gt;&quot; &quot;&lt;&quot; &quot;&lt;&gt;&quot; 
  &quot;&gt;=&quot; &quot;&lt;=&quot; &quot;=&quot; ��ô�q�I<br>
  �b�d�ߪ��~�ɡG�d�ߦr�q���ȵL�ġA�p��J�G�֦���='�����`��'<br>
  <font color="#0000FF">�d�ߦr�q�G���Ȭ���@��,�p�G���O�B��O���C�H�U�����~�G���O and ��O �� ���O,��O��</font></p>
<div align="center">�Τ��ƭק�{�� </div>
<form method="POST" action="seeuserok.asp">
  <div align="center"><font color="#FFFFFF">
    <select name="seekfs" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="0" selected>�d�ߥΤ�ƾڮw</option>
      <option value="1">�d�ߪ��~�ƾڮw</option>
      <option value="2">�d�߭׷Ҫ��~�w</option>
    </select>
    </font><br>
    <br>
    �п�J�d�߱���G 
    <input type="text" name="tiaojian" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="50" maxlength="250">
    <br>
    ��J�N�d�ߦr�q�G 
    <input type="text" name="show" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="12">
    <br>
    <input type="submit" value="�d��" name="B1" class="p9">
    <input type="reset" name="Submit" value="�M��">
  </div>
  </form>
</body>
</html> 