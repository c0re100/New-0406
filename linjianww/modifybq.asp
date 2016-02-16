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

wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 物品買 where id="&wupinid
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Redirect "modifywupin.htm"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font size="+6" color="#FF0000">物品修改</font><font size="2"><br>
  <br>
  注：類型修改請填藥品，暗器，兵器，頭盔，盔甲，雙腳，裝飾，毒藥,卡片九種類型，否則添加無效。</font></p>
<form method="post" action="modifybqnow.asp">
  <table border="1" cellspacing="1" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="305" bgcolor="#FFFFFF" cellpadding="1">
    <tr bgcolor="#99CCFF"> 
      <td width="105"><font size="2">ID</font></td>
      <td width="189" bgcolor="#99CCFF"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinid" readonly
value="<%=rst("ID")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">物品名</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinname"
value="<%=rst("物品名")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">類型</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinlx"
value="<%=rst("類型")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">內力</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinnl"
value="<%=rst("內力")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">體力</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupintl"
value="<%=rst("體力")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">攻擊</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupingj"
value="<%=rst("攻擊")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">防御</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinfy"
value="<%=rst("防御")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">堅固度</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinjg"
value="<%=rst("堅固度")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">說明</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinsm"
value="<%=rst("說明")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">等級</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupindj"
value="<%=rst("等級")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">銀兩</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinyl"
value="<%=rst("銀兩")%>" size="20">
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
<%
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<div align="center"> 
  <p><font size="2">數據庫更新程序因為時間有限，沒有加入為空值時的檢測，請大家更改時注意沒有值的地方填0</font> </p>
</div>
