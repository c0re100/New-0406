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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<20 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=20-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');window.close();}</script>"
	Response.End
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
tx="../ico/"& info(10) &"-2.gif"

action=Request.Querystring("action")
if action="attack" then
Randomize timer
points1=int(100*rnd)+1
points2=int(100*rnd)+1
r=rnd+rnd+1
if points1=points2 then points2=points2+1
Response.Cookies("points1")=points1
Response.Cookies("points2")=points2
Response.Cookies("attack")=""
Response.write "<script language='javascript'>alert('會有什麼事發生呢？');location.href='gs.asp?"& r &"';</script>"
else
hp1=Request.Cookies("hp1")
hp2=Request.Cookies("hp2")
points1=Request.Cookies("points1")
points2=Request.Cookies("points2")
attack=Request.Cookies("attack")
if hp1="" then hp1=100
if hp2="" then hp2=100
if hp1<=0 then
   response.cookies("hp1")=""
   response.cookies("hp2")=""
   Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
   conn.execute("update 用戶 set 銀兩=銀兩+1000000,操作時間=getdate() where id="&info(9))
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=red>【怪獸消息】["&info(0)&"]</font>挑戰怪獸，經過一番激鬥成功打倒怪獸獲得賞銀100萬!(想賺錢就來這)" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
   response.write "<script language='javascript'>alert('怪獸被打倒了，獲得銀兩100萬！！！');window.close();</script>"
end if
if hp2<=0 then
   response.cookies("hp1")=""
   response.cookies("hp2")=""
   response.write "<script language='javascript'>alert('我被打倒了！！！');window.close();</script>"
end if
%>
<html>
<head>
<title>挑戰怪獸家族</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px; line-height: 22px}
input {  font-family: "新細明體"; font-size: 12px; background-color: #FFFFCC; color: #660000; border: #660000; border-style: dotted; border-top-width: thin; border-right-width: thin; border-bottom-width: thin; border-left-width: thin}
a:link {  font-family: "新細明體"; font-size: 12px; color: #660000; text-decoration: none}
a:visited {  font-family: "新細明體"; font-size: 12px; color: #660000; text-decoration: none}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" link="#660000" vlink="#660000" background="../ljimage/11.jpg">
<br>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td height="25">挑戰怪獸家族 （勝利後獲100萬兩，太簡單了吧！！！）</td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td width="30%" height="37">怪獸<img src="../IMAGES/010.gif" width="79" height="43"> 
      生命點 
      <input type="text" readonly size="4" value="<%=hp1%>">
  </td>
    <td width="40%" height="37">&nbsp;</td>
    <td width="30%" height="37">獵殺者<img src="<%=tx%>" >生命點 
      <input type="text" readonly size="4" value="<%=hp2%>">
    </td>
  </tr>
  <tr align="center"> 
    <td height="64">&nbsp;</td>
    <td height="64"> <%
Randomize timer
r=rnd+rnd+1
if (points1<>"" or points2<>"") and attack="" then
      if points1<points2 then
         response.cookies("hp1")=hp1-10
         response.cookies("attack")="yes"
         response.write "怪物怒氣：<font color=red>"&points1&"</font> VS 獵殺者怒氣：<font color=red>"&points2&"</font><br><a href='gs.asp?"& r &"'>點擊刷新生命點</a>"
      else
         response.cookies("hp2")=hp2-10
         response.cookies("attack")="yes"
         response.write "怪物怒氣：<font color=red>"&points1&"</font> VS 獵殺者怒氣：<font color=red>"&points2&"</font><br><a href='gs.asp?"& r &"'>點擊刷新生命點</a>"
      end if
end if
%> </td>
    <td height="64">
      <input type="button" name="Button" value="點數攻擊" OnClick="javasript:location.href='gs.asp?action=attack&<%=r%>';">
      <input type="button" name="Button" value="退出戰場" OnClick="javascript:window.close();">
    </td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td height="40">本插件全部使用瀏覽器的 Cookie 功能，請不要關閉此功能<br>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>