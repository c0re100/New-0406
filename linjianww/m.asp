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
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>�K�[�s�������I</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="0064.jpg">
<form action="newzt.asp" method=POST >
  <ul>
    <table border=0 cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="49"> 
          <p align="center"><font size="2" class=c><br>
            </font><b><font  size="+6" color="#FF0000">�����K�[</font></b><font size="2" class=c><b><br>
            </b> <br>
            </font></p>
        </td>
      </tr>
      <tr> 
        <td height="2"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">�����W:</font></font> 
            <input name="subject" value="" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=10>
            <br>
            <br>
            �`�G�����W�ٳ̦h��10�ӭ^��w�šA5�Ӥ���r��<br>
            <br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">² 
            ��:</font></font> 
            <input name="comment" value="" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=30>
            <br>
            <br>
            �`�G����²���̦h��30�Ӧr�šA�B���i���šI<br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1"><br>
            �̤l�ʧO:</font></font> 
            <select name="ljxb" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
              <option value="�k">�k</option>
              <option value="�k">�k</option>
              <option value="�Ҧ��H" selected>�Ҧ��H</option>
            </select>
            <br>
            <br>
            �`�G�̤l�ʧO���Ӫ��������\�k/�k/�Ҧ��H���[�J�I</div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="�`�U">
      <input type="RESET" name="Reset" value="�M��">
      </font> </div>
  </ul>
</form>
</body>
</html>
