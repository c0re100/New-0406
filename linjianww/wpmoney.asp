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
%>
<html>
<head><style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<title>���~�R����]�㾹</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if(document.FrmAddLink.moneybei.value=="")
{
alert("�Ȩ⭿�ƨS���g�A�L�k�����{�ǡI")
document.FrmAddLink.moneybei.focus()
return false
}
}

//-->
</script>
</head>
<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method=post action="wpmoney1.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
  <div align="center"><font size="+6" color="#FF0000">���~�R����]�㾹</font> </div>
  <p>&nbsp;</p><table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" width="229">
    <tr> 
      <td width="66"> 
        <div align="center"><font color="#FF6600" size="2">����</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <select name="lx" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <option value="�k��" selected>����M�C</option>
          <option value="����">����@��</option>
          <option value="�t��">�t��</option>
          <option value="�Y��">�Y��</option>
          <option value="���}">���}</option>
          <option value="����">����</option>
          <option value="�˹�">�˹�</option>
          <option value="�ī~">�����</option>
          <option value="�r��">�r��</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="66"> 
        <div align="center"><font color="#FF6600" size="2">�Ȩ⭿��</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=moneybei style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="0">
        </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type=submit value="�T �w" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='binqi.asp'" value="�� �^" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<div align="center">
  <p><br>
    �`�G��ܸ˳������A��ܻȨ⭿�ơA�T�U�N�i�H�۰ʧ�s���~�����I<br>
    ���~����= (���O+��O+����+���s)*�Ȩ⭿��<br>
    �p�G���B��B��B�����X�{�t�ơA�N�۰��ഫ�����ơI</p>
  <p><br>
</div>
</body>
</html> 