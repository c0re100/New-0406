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
<title>裝備添加</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.wupinname.value=="")
{
alert("名稱沒有填！")
document.FrmAddLink.wupinname.focus()
return false
}
else if(document.FrmAddLink.wupinll.value=="")
{
alert("增加攻擊沒有填！")
document.FrmAddLink.wupinll.focus()
return false
}
else if(document.FrmAddLink.wupintl.value=="")
{
alert("增加防御沒有填！")
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
        <div align="center"><font color="#FFFFFF" size="2">兵器名</font></div>
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
          <option value="右手" selected>手持刀劍</option>
          <option value="左手">手持護具</option>
          <option value="暗器">暗器</option>
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
      <td width="57"> 
        <div align="center"><font color="#FFFFFF" size="2">等 級</font></div>
      </td>
      <td width="154"><font color="#FFFFFF" size="2"> 
        <input type="text" name=wupindj value="0">
        </font></td>
    </tr>
    <tr> 
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
</body>
</html> 