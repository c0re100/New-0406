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
%>
<html>
<head>

<title>���~�o�e�޲z</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[���~�o�e�޲z]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">�m�W</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">���~</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">�ƶq</font></div>
    </td>
    <td height="27" width="250"> 
      <div align="center"><font color="#FF6600">�ާ@</font></div>
    </td>
  </tr>
    <form method=POST action='fawupinok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12" maxlength="12">
        </div>
      </td>
      <td height="2" width="142"> 
        <div align="center"> 
          <select name="b">
            <option value="���C" selected>���C</option>
            <option value="�����C">�����C</option>
            <option value="�Q�ڼC">�Q�ڼC</option>
            <option value="���Z�C">���Z�C</option>
            <option value="�ȭ߼C">�ȭ߼C</option>
            <option value="�H�P�]�C">�H�P�]�C</option>
            <option value="�涳���b">�涳���b</option>
            <option value="�ݦ�C">�ݦ�C</option>
            <option value="�C�Y�]�C">�C�Y�]�C</option>
            <option value="�ĳ��C">�ĳ��C</option>
            <option value="���s�C">���s�C</option>
            <option value="�C�P�_�C">�C�P�_�C</option>
            <option value="�������M">�������M</option>
            <option value="�񲴼C">�񲴼C</option>
            <option value="���ѼC">���ѼC</option>
            <option value="�ܴL�C">�ܴL�C</option>
            <option value="�O�s�M">�O�s�M</option>
            <option value="�s�]">�s�]</option>
          </select>
        </div>
      </td>
      <td height="2" width="142"> 
        <div align="center"> 
          <select name="c">
            <option value="1" selected>1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
          </select>
        </div>
      </td>
      <td height="2" width="150"> 
        <div align="center"> 
          <input type="submit" value="�o��"
name="submit">
          &nbsp;&nbsp; </div>
      </td>
    </tr>
  </form>

</table>
</body>
</html>