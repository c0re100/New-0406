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
<title>藥品添加</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.wupinname.value=="")
{
alert("物品名稱沒有填！")
document.FrmAddLink.wupinname.focus()
return false
}
else if(document.FrmAddLink.wupinll.value=="")
{
alert("增加內力沒有填！")
document.FrmAddLink.wupinll.focus()
return false
}
else if(document.FrmAddLink.wupintl.value=="")
{
alert("增加體力沒有填！")
document.FrmAddLink.wupintl.focus()
return false
}
}

//-->
</script>
</head>
<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method=post action="yaoaddok.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
  <table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="208">
    <tr> 
      <td width="57" height="29"> 
        <div align="center"><font color="#FFFFFF" size="2">藥名</font></div>
      </td>
      <td width="154" height="29"><font color="#FFFFFF" size="2"> 
        <INPUT type="text" name=wupinname>
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">類型</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <select name="lx">
          <option value="藥品" selected>藥品</option>
          <option value="毒藥">毒藥</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">防御</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinfy value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">攻擊</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupingj value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">體力</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupintl value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">內力</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinnl value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">堅固度</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinjg value="0">
        </font></td>
    </tr>
    <tr> 
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">說 明</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupinsay value="0">
        </font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type=submit value="確 定" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='yaopu.asp'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>