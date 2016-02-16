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
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#FF0000" size="+6">添加卡片</font><FONT COLOR="#FFFFFF" SIZE="2"><br>
  <br>
  </FONT><FONT SIZE="2">注：類型卡片不要動，否則添加無效。</FONT></p>
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
      <td width="105"><FONT SIZE="2">卡片名</FONT></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinname" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">類型</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinlx" size="20" VALUE="卡片">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">內力</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinnl" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">體力</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupintl" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">攻擊</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupingj" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">防御</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinfy" size="20" VALUE="1">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">說明</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinsm" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><FONT SIZE="2">價格</FONT></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinyl" size="20">
        </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="確 定">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='javascript:history.go(-1)'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
