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
<p align="left">�޲z�����G<br>
  <font color="yellow"><br>
  </font>�{�b���@�Ǧr�q�S�����G�p����B����[�B�y�O�B�y�O�[�B�|���O���o�ǬO���F�H��\��ɯŮɨϥΪ��A�O�@�ǳƥΪ��C<font color="yellow"><br>
  </font><font color="#0000FF">�|���ɶ��O�I��ɶ��A�|�����Ŭ��G1,2,3,4�p�G��Ȭ�0�A�h��ܤ��O�|���I</font><br>
  �Z�\�[�A���O�[�A��O�[�A�����[���C�o�@�ǭȬO���n�ç諸�A�o�ǬO�L�̦b����W�����B���رo�쪺�o�ǭȤ��n��ʡI�p�G�b�޲z�W�����򤣩��ժ����D�A�Ч��F�C���� 
  �q�l:Seven.s-888@yahoo.com.tw</p>
<div align="center">�Τ��ƭק�{�� </div>
<form method="POST" action="showuser.asp">
  <div align="center"> 
    <select name="sjcz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="�m�W" selected>�m�W</option>
      <option value="ID">ID</option>
      <option value="����">����</option>
      <option value="�t��">�t��</option>
      <option value="�H�c">�H�c</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="10">
    <input type="submit" value="�d��" name="B1" class="p9">
    <input type="reset" name="Submit" value="�M��">
  </div>
  <div align="center">ID�d��@�w�n���Ʀr�I<br>
  </div>
  </form>
<form method="POST" action="laren.asp">
  <div align="center">�ԤH�d�ݵ{�ǡI<br>
    <input type="text" name="larenseek" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="�d��" name="B12" class="p9">
    <input type="reset" name="Submit2" value="�M��">
  </div>
  <div align="center"></div>
</form>
<form method="POST" action="pass.asp">
  <div align="center">�K�X�ץ��{��<br>��J�H�W�N�L���K�X�令�G123456<br>
    <input type="text" name="cpass" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"  maxlength="10">
    <input type="submit" value="�T�w" name="B12" class="p9">
    <input type="reset" name="Submit2" value="�M��">
    <br>
  </div>
  </form>
<form method="POST" action="jiaji.asp">
  <div align="center">���a�԰��ť[�ŵ{��<br>
    ��J�H�W�M�ҭn�[���ż�<br>
    <input type="text" name="cpass2" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="text" name="cpass22" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="�T�w" name="B122" class="p9">
    <input type="reset" name="Submit22" value="�M��">
    <br>
  </div>
</form>
<form method="POST" action="jieji.asp">
  <div align="center">���a�԰��Ŵ�ŵ{��<br>
    ��J�H�W�M�ҭn��ż�<br>
    <input type="text" name="cpass23" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="text" name="cpass222" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="�T�w" name="B1222" class="p9">
    <input type="reset" name="Submit222" value="�M��">
    <br>
  </div>
</form>
</body>
</html>
 