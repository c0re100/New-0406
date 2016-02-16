<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
%><!--#include file="data.asp"--><%
'rs.open "select * from 用戶 where id="&info(9),conn
set rs1=conn1.execute("select house,bigarea from userinfo where user='"&info(0)&"'")
If Rs1.Bof OR Rs1.Eof Then
hh="0"
conn1.execute"insert into userinfo(user,house) values('"&info(0)&"','"&hh&"')"
end if
rs1.close
%>
<html>

<head>
<title>靈劍江湖-房屋交易所</title>
<link rel="stylesheet" type="text/css" href="../style.css">
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

<body leftmargin="0" topmargin="7" bgcolor="#FFFFFF">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../images/fw.gif" width="101" height="304"></td>
<td colspan="2" valign="top" align="center">
<table width="61%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="90" height="26">&nbsp;<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %> </td>
<td width="302" height="26" align="center"><font color="#FF6600">房 屋
交 易 所</font></td>
<td width="73">
<div align="center"> <a href="javascript:history.back()">返回</a></div>
</td>
</tr>
</table>
<br>
<form method="POST" action="housedispose.asp">
<div align="center">
<center>
<table border="0" cellspacing="3" cellpadding="0" width="650">
<tr>
<td class="p3" colspan="3">
<input
type="radio" name="S1" value="buy" checked>
□- 買
<input
type="radio" name="S1" value="sale">
□- 賣
<select size="1"
name="area">
<option value="夏龍">夏龍區</option>
<option value="雲界">雲界區</option>
<option value="夢縈">夢縈區</option>
<option value="天穹">天穹區</option>
<option value="財緣">財緣區</option>
<option value="芹獻">芹獻區</option>
</select>
</td>
</tr>
<tr>
<td class="p1" colspan="3"><font color="#0000FF">□-請根據您的興趣愛好和生活目標選擇街區！</font></td>
</tr>
<tr>
<td class="p3" colspan="2">□-夏龍區</td>
<td class="p3" width="552">華夏之龍,您是我們華夏從古到今的英雄人物。</td>
</tr>
<tr>
<td class="p2" colspan="2">□-雲界區</td>
<td class="p2" width="552">凌雲之界,您是別人所比不上的奇纔,怪纔,天纔。</td>
</tr>
<tr>
<td class="p3" colspan="2">□-夢縈區</td>
<td class="p3" width="552">夢裡依稀,您喜歡幻想,您是童話神話中的人。</td>
</tr>
<tr>
<td class="p2" colspan="2">□-天穹區</td>
<td class="p2" width="552">天馬行空,您喜歡科幻,遊戲是縱橫銀河的無敵霸者。</td>
</tr>
<tr>
<td class="p3" colspan="2">□-財緣區</td>
<td class="p3" width="552">財&quot;緣&quot;滾滾,您是百裡挑一的精商,&quot;奸商&quot;(?)。</td>
</tr>
<tr>
<td class="p2" colspan="2">□-芹獻區</td>
<td class="p2" width="552">&quot;勤&quot;勞奉獻,您是默默無聞的人,但城市不能少了你們。</td>
</tr>
<tr>
<td class="p1" colspan="3"><font color="#FF0000">□-請根據您的經濟情況選擇下面的房屋：）--我們可不賒帳的哦
！</font></td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house1" name="R1" checked>
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h01.gif"></td>
<td class="p2" valign="top" width="552">□- [簡陋茅屋] 用蘆葦、稻草等苫蓋屋頂的簡陋房子，是貧民窟中最典型的房子。<br>
價格-[<font color="#FF0000">800</font>元]</td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house2" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h02.gif"></td>
<td class="p3" valign="top" width="552">□- [一般平房] 只有一層的小瓦房，很多剛從貧民窟搬出來的人住在這裡。<br>
價格-[<font color="#FF0000">20000</font>元]</td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house3" name="R1">
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h03.gif"></td>
<td class="p2" width="552">□- [普通公寓] 多少是按永久存在設計而建成的建築物,占用土地空間,通常有屋頂,多半完全用牆包圍住,作為住宅、倉庫、工廠、牲畜圈棚或其他有用的建築物。價格-[<font
color="#FF0000">50000</font>元]</td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house4" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h04.gif"></td>
<td class="p3" valign="top" width="552">□- [豪華公寓] 皇城的中層收入者都住在這裡，大多是各門派有點職位的或運氣比較好的探險者。&nbsp;
價格-[<font color="#FF0000">200000</font>元]</td>
</tr>
<tr>
<td class="p2" width="24">
<input type="radio"
value="house5" name="R1">
</td>
<td class="p2" width="59"><img border="0"
src="../jhimg/h05.gif"></td>
<td class="p2" valign="top" width="552">□- [普通別墅] 修在黃金地段的別墅，交通方便，環境優美。住在這裡的都大多是各派掌門或商店的老板哦！價格-[<font
color="#FF0000">500000</font>元] </td>
</tr>
<tr>
<td class="p3" width="24">
<input type="radio"
value="house6" name="R1">
</td>
<td class="p3" width="59"><img border="0"
src="../jhimg/h06.gif"></td>
<td class="p3" valign="top" width="552">□- [豪華別墅]-在風景區或在郊區建造的供休養的豪華別墅，皇城的長老、官員和大富翁們住在這裡。&nbsp;&nbsp;
價格-[<font color="#FF0000">1000000</font>元] </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="468" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="462" height="13">
<p align="center">
<input type="submit" value="決定"
name="B1" style="font-size: 9pt">
<input type="reset"
value="考慮一下" name="B2" style="font-size: 9pt">
</td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>

</body>

</html>
