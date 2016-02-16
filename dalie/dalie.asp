<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	response.end
end if
if session("dalie")=false then
	Response.Write "<script Language=Javascript>alert('提示：想作弊嗎？！');window.close();</script>"
	response.end
end if
if Application("ljjh_dalie")<>"老虎" then
	Response.Write "<script Language=Javascript>alert('提示：現在還沒有獵物可以打，請一會再來吧。');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_dalie")="無"
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 內力=內力+800,體力=體力+1200,allvalue=allvalue+1,銀兩=銀兩+5000 where id="&info(9)
rs.open "select id from 物品 where 物品名='老虎肉' and 擁有者=" &info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('老虎肉',"&info(9)&",'藥品',0,0,0,150,150,1,2000,400,400,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 銀兩=2000,數量=數量+1,會員="&aaao&" where 物品名='老虎肉' and id="&id
end if
mess="<b><font color=red>野獸已經被["&info(0)&"]打死，真是英雄少年呀，請大家放心聊天吧!</b>"
randomize timer
r=int(rnd*3)+1
if r=3 then 
rs.close
rs.open "select id from 物品 where 物品名='明月秘訣' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('明月秘訣',"&info(9)&",'物品',0,0,0,0,0,1,200000,99003,99003,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
mess="<b><font color=red>野獸已經被["&info(0)&"]打死，真是英雄少年呀，請大家放心聊天吧,"&info(0)&"在野獸窩找到一本<font color=FFFDAF>《明月秘訣》<img src='../hcjs/jhjs/001/99003.gif' border='0'></font></font></b>"
end if
rs.close
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
sd(199)="<font color=FFFDAF>【好消息】"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<title>打獵成功</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>

<body text="#000000" background="../../images/8.jpg"" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#0000FF" align="center"><font color="#FFFFFF">打獵成功</font></td>
</tr>
<tr>
<td bgcolor="#FFCC99" align="center" height="157"><font> 您成功的將野獸打死，增加內力+800，體力+1200，銀兩+5000，點券+1。<br>
<br>
虎肉一塊，使用物品:內力+150,體力+150</font></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>
