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

%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#FF0000" size="+6">�K�[�d��</font><FONT COLOR="#FFFFFF" SIZE="2"><br>
  <br>
  </FONT><FONT SIZE="2">�`�G�����d�����n�ʡA�_�h�K�[�L�ġC</FONT></p>
<form method="post" action="addcardok.asp"> 
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"
bgcolor="#FFFFFF" width="305">
    <tr bgcolor="#99CCFF"> 
      <td width="105"><font size="2">ID</font></td>
      <td width="189" bgcolor="#99CCFF"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinid" readonly size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><FONT SIZE="2">�d���W</FONT></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinname" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinlx" size="20" VALUE="�d��">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">���O</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinnl" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">��O</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupintl" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupingj" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">���s</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinfy" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinsm" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><FONT SIZE="2">����</FONT></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinyl" size="20">
        </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="�T �w">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='javascript:history.go(-1)'" value="�� �^" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
