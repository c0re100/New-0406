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

<title>�R���p�κ޲z</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+3">[�R���p�κ޲z]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="600"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="85"> 
      <div align="center"><font color="#FF6600"> �� ��</font></div>
    </td>
    <td height="27" width="80"> 
      <div align="center"><font color="#FF6600"> �a �� </font></div>
    </td>
    <td height="27" width="84" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600"> �� �� </font></div>
    </td>
    <td height="27" width="88"> 
      <div align="center"><font color="#FF6600"> �W�[�Фl�� </font></div>
    </td>
    <td height="27" width="81"> 
      <div align="center"><font color="#FF6600"> �� �� �� �@ </font></div>
    </td>
  </tr>
  
  <%

if s<50 then
%>
  <form method=POST action='fangwuok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="85"> 
        <div align="center"> 
         <select name="leixin" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="�@��Ы�" selected>�@��Ы�</option>
      <option value="���Ť��J">���Ť��J</option>
      <option value="���v��">���v��</option>
      <option value="���اO��">���اO��</option>
    </select>
        </div>
      </td>
      <td height="2" width="160"> 
        <div align="center"> 
          <select name="jiedao" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="������" selected>������</option>
      <option value="�بʰ�">�بʰ�</option>
      <option value="�ص���">�ص���</option>
      <option value="�ضQ��">�ضQ��</option>
    </select>
        </div>
      </td>
<td height="2" width="160"> 
        <div align="center"> 
          <select name="money" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="500000" selected>500000</option>
      <option value="1000000">1000000</option>
      <option value="3000000">3000000</option>
      <option value="5000000">5000000</option>
    </select>
        </div>
      </td>
      <td width="154" height="2"> 
        <div align="center"> 
         <select name="sn" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="1" selected>1</option>
      <option value="5">5</option>
      <option value="10">10</option>
      <option value="15">15</option>
      <option value="20">20</option>
    </select>
        </div>
      </td>
     
      <td width="141" height="2"> 
        <div align="center"> 
          <input type="submit" value="�K�["
name="submit">
        </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center">�H�W�Ů�B���ର��<br>
</body>
</html>