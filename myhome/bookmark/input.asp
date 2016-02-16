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

 
<!--#include file="../data1.asp"-->
<html>

<head>
<title>靈劍江湖-寵物商店</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000" text="#FFFFFF">

<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
<td width="744" background="../../jh/tiao20.gif" height="83">
<table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"><span class="zilong"><font color="#FFCC00">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %>
</font></span></font></font></font></td>
<td width="475" height="26" align="center"><font size="3"
color="#E18C59"><b>俠客寓所--&nbsp;書簽</b></font></td>
<td width="104">
<div align="center">

</div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
</td>
</tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="17" background="../../jh/tiao10.gif">&nbsp;</td>
<td width="47" valign="top">
<div align="center">
<img src="../../jh/17.gif" width="47" height="251">
</div>
</td>
<td valign="top" width="639">
<div align="center">
<div align="center">
<center>
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">
</center>
</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="635" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">□-您當前的位置-俠客寓所--&nbsp;我的書簽&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;<a href="javascript:history.back(1)"><font color="#FFCC00">返回</font></a></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<form method="POST" action="post.asp">
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">網站名稱：</td>
<td class="p3" width="281"><input type="text" name="name"
size="32"></td>
<td class="p3" width="100">具體分類：</td>
<td class="p3" width="265"><input type="text" name="emptytype"
size="32"></td>
</tr>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">選擇分類：</td>
<td class="p3" width="281"><select size="1" name="type">
<option value="娛樂休閑">娛樂休閑</option>
<option value="商業經濟">商業經濟</option>
<option value="電腦網絡">電腦網絡</option>
<option value="醫療健康">醫療健康</option>
<option value="文學藝術">文學藝術</option>
<option value="教育就業">教育就業</option>
<option value="科學技術">科學技術</option>
<option value="體育健身">體育健身</option>
<option value="社會文化">社會文化</option>
<option value="旅遊交通">旅遊交通</option>
<option value="政法軍事">政法軍事</option>
<option value="生活服務">生活服務</option>
<option value="社會科學">社會科學</option>
<option value="參考資料">參考資料</option>
<option value="新聞媒體">新聞媒體</option>
<option value="遊戲天地">遊戲天地</option>
<option value="音樂網站">音樂網站</option>
<option value="搜索引擎">搜索引擎</option>
<option value="圖書網站">圖書網站</option>
<option value="軟件下載">軟件下載</option>
<option value="電子郵箱">電子郵箱</option>
<option value="免費資源">免費資源</option>
<option value="圖庫網站">圖庫網站</option>
<option value="硬件天地">硬件天地</option>
<option value="聊天交友">聊天交友</option>
<option value="網絡編程">網絡編程</option>
<option value="社區論壇">社區論壇</option>
<option value="網上教學">網上教學</option>
<option value="電子賀卡">電子賀卡</option>
</select></td>
<td class="p3" width="100">網站地址：</td>
<td class="p3" width="265">http://<input type="text" name="url"
size="27"></td>
</tr>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="114">是否開放：</td>
<td class="p3" width="281"><input type="radio" value="是"
checked name="open">是<input type="radio" name="open"
value="否">否</td>
<td class="p3" width="100"></td>
<td class="p3" width="265"><input type="submit" value="提交"
name="B1" style="font-family: Tahoma; font-size: 12px"><input
type="reset" value="全部重寫" name="B2"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</form>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">&nbsp;</td>
</tr>
</table>
</center>
</div>
</div>
</td>
<td width="25" background="../../jh/tiao10.gif"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<div align="right">
<img src="../../jh/tiao19.gif" width="728" height="31">
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>