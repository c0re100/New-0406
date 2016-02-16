<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
%>
<!--#include file="homecheck.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<title><%=username%>你回家了！</title>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"
bgcolor="#000000" text="#000000" background="../ljimage/11.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table border="0" height="17" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td height="26">
<div align="center"><font color="#FF0000">  <b><font size="2">你回到了自己的家，這裡沒有危險，你可以好好休息一下了！</font></b>
</font></div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
<table width="738" border="1" cellspacing="0" cellpadding="0" align="center" height="507" bordercolor="#000000">
<tr>
<td width="145" valign="top" height="351"><font size="2" color="#000000">當前時間：<%=hour(now)&":"&minute(now)&":"&second(now)%><br>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select 姓名,性別,配偶,門派,身份,武功,內力,體力,攻擊,防御,銀兩,狀態,道德,等級,魅力 from 用戶 where 姓名='"&info(0)&"'")%>
歡迎<%=rs("姓名")%>回家<br>
性別:<%=rs("性別")%><br>
配偶:<%=rs("配偶")%><br>
門派:<%=rs("門派")%><br>
身份:<%=rs("身份")%><br>
武功:<%=rs("武功")%><br>
內力:<%=rs("內力")%><br>
體力:<%=rs("體力")%><br>
攻擊:<%=rs("攻擊")%><br>
防御:<%=rs("防御")%><br>
銀兩:<%=rs("銀兩")%><br>
狀態:<%=rs("狀態")%><br>
道德:<%=rs("道德")%><br>
等級:<%=rs("等級")%><br>
魅力:<%=rs("魅力")%><br>
<%set rs=nothing
conn.close%>
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select house,bigarea,homeopen from userinfo where user='"&info(0)&"'")%>
房屋類型:
<%if rs("house")="house1" then%>
破爛茅屋
<%end if%>
<%if rs("house")="house2" then%>
一般平房
<%end if%>
<%if rs("house")="house3" then%>
一般公寓
<%end if%>
<%if rs("house")="house4" then%>
豪華公寓
<%end if%>
<%if rs("house")="house5" then%>
普通別墅
<%end if%>
<%if rs("house")="house6" then%>
豪華別墅
<%end if%>
<br>
所在區:<%=rs("bigarea")%>區<br>
家門狀態：
<%if rs("homeopen")="1" then%>
大門敞開著
<%else%>
大門緊閉著</font><font color="#000000">
<%end if%>
</font></td>
<td valign="top" height="351">
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" width="100%"
height="307" align="center">
<tr>
<td align="center" colspan="2" valign="top" width="648" height="451">
<div align="left">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
height="21" align="center" bordercolorlight="#000000"
bordercolordark="#FFFFFF">
<tr align="center" valign="middle">
<td height="6" align="left"><font size="2" color="#000000">你回到位於
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select bigarea,house from userinfo where user='"&info(0)&"'")%>
<%=rs("bigarea")%>區的
<%if rs("house")="house1" then%>
破爛小茅屋裡，這是你的家。你的經濟狀態不是很好！要多努力哦！當然除工作也要注意身體！
<%end if%>
<%if rs("house")="house2" then%>
小平房裡，這裡是你的家。你在靈劍江湖立足為久，在朋友們的幫助下蓋起了這個小屋！
<%end if%>
<%if rs("house")="house3" then%>
一般公寓裡，這裡是你的家。你是靈劍江湖裡的一般俠客，有一定的收入~相信經過你的努力，不久就可以搬到更大的房子裡了！
<%end if%>
<%if rs("house")="house4" then%>
豪華公寓裡，這裡是你的家。這是一所漂亮的房子，四處透露出家的溫暖。你在靈劍江湖混得不錯~~繼續努力吧！
<%end if%>
<%if rs("house")="house5" then%>
普通別墅裡，這裡是你的家。別墅四周風景宜人~僕人們為你端茶到水！你是靈劍江湖裡的實力派，誰也不敢隨便惹你！
<%end if%>
<%if rs("house")="house6" then%>
豪華別墅裡，這裡是您的府邸。管家帶著大隊的人</font>
<div id="Layer2"
style="position:absolute; width:200px; height:115px; z-index:2; left: 448px; top: 159px">
<font size="2" color="#000000"><img src="christmas1.jpg"
width="279" height="181"></font> </div>
<font size="2" color="#000000">馬把您迎接進了客廳，為您送上了名貴清茶！您看著經過努力而得到的一切，心中感慨萬千！您是靈劍江湖裡的靈劍江湖人物！</font>
<font color="#000000">
<%end if%>
</font></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
</table>
<div id="Layer1"
style="position:absolute; width:119px; height:91px; z-index:1; left: 241px; top: 199px">
<font color="#000000"><img src="a1.gif" width="118" height="88">
</font></div>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</body>
</html>
