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

<title>ª««~µo°eºÞ²z</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: ·s²Ó©úÅé}
.c {  font-family: ·s²Ó©úÅé; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[ª««~µo°eºÞ²z]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">©m¦W</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">ª««~</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">¼Æ¶q</font></div>
    </td>
    <td height="27" width="250"> 
      <div align="center"><font color="#FF6600">¾Þ§@</font></div>
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
            <option value="¨ª¼C" selected>¨ª¼C</option>
            <option value="½¹½º¼C">½¹½º¼C</option>
            <option value="ªQ®Ú¼C">ªQ®Ú¼C</option>
            <option value="¨ëÆZ¼C">¨ëÆZ¼C</option>
            <option value="¥È­ß¼C">¥È­ß¼C</option>
            <option value="´H¬PÅ]¼C">´H¬PÅ]¼C</option>
            <option value="¦æ¶³°­¤b">¦æ¶³°­¤b</option>
            <option value="¶Ý¦å¼C">¶Ý¦å¼C</option>
            <option value="«CÀYÅ]¼C">«CÀYÅ]¼C</option>
            <option value="¿Ä³·¼C">¿Ä³·¼C</option>
            <option value="µëÀs¼C">µëÀs¼C</option>
            <option value="¤C¬PÄ_¼C">¤C¬PÄ_¼C</option>
            <option value="µ²«ô¥¨¤M">µ²«ô¥¨¤M</option>
            <option value="»ñ²´¼C">»ñ²´¼C</option>
            <option value="¨ªºÑ¼C">¨ªºÑ¼C</option>
            <option value="¦Ü´L¼C">¦Ü´L¼C</option>
            <option value="±OÀs¤M">±OÀs¤M</option>
            <option value="Às¯]">Às¯]</option>
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
          <input type="submit" value="µo©ñ"
name="submit">
          &nbsp;&nbsp; </div>
      </td>
    </tr>
  </form>

</table>
</body>
</html>