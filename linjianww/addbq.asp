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
<%
Response.Expires=0
Response.Buffer=true
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
%>
<html>
<head>
<title>�˳ƲK�[</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.wupinname.value=="")
{
alert("�W�٨S����I")
document.FrmAddLink.wupinname.focus()
return false
}
else if(document.FrmAddLink.wupinll.value=="")
{
alert("�W�[�����S����I")
document.FrmAddLink.wupinll.focus()
return false
}
else if(document.FrmAddLink.wupintl.value=="")
{
alert("�W�[���s�S����I")
document.FrmAddLink.wupintl.focus()
return false
}
}

//-->
</script>
</head>
<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method=post action="bqaddok.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
  <table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="208">
    <tr> 
      <td width="57" height="29"> 
        <div align="center"><font color="#FFFFFF" size="2">�L���W</font></div>
      </td>
      <td width="154" height="29"><font color="#FFFFFF" size="2"> 
        <INPUT type="text" name=wupinname>
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <select name="lx">
          <option value="�k��" selected>����M�C</option>
          <option value="����">����@��</option>
          <option value="�t��">�t��</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">���s</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinfy value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupingj value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">��O</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupintl value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">���O</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinnl value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">��T��</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinjg value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">�� ��</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinsay value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">�� ��</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupindj value="0">
        </font></td>
    </tr>
    <tr> 
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
</body>
</html> 