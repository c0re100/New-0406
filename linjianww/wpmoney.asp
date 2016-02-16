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
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<title>物品買價格設整器</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if(document.FrmAddLink.moneybei.value=="")
{
alert("銀兩倍數沒有寫，無法完成程序！")
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
  <div align="center"><font size="+6" color="#FF0000">物品買價格設整器</font> </div>
  <p>&nbsp;</p><table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" width="229">
    <tr> 
      <td width="66"> 
        <div align="center"><font color="#FF6600" size="2">類型</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <select name="lx" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <option value="右手" selected>手持刀劍</option>
          <option value="左手">手持護具</option>
          <option value="暗器">暗器</option>
          <option value="頭盔">頭盔</option>
          <option value="雙腳">雙腳</option>
          <option value="盔甲">盔甲</option>
          <option value="裝飾">裝飾</option>
          <option value="藥品">喫用藥</option>
          <option value="毒藥">毒藥</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="66"> 
        <div align="center"><font color="#FF6600" size="2">銀兩倍數</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=moneybei style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="0">
        </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type=submit value="確 定" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='binqi.asp'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<div align="center">
  <p><br>
    注：選擇裝備類型，選擇銀兩倍數，確下就可以自動更新物品價錢！<br>
    物品價錢= (內力+體力+攻擊+防御)*銀兩倍數<br>
    如果內、體、攻、防等出現負數，將自動轉換成正數！</p>
  <p><br>
</div>
</body>
</html> 