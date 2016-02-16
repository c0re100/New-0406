<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
abc=DateDiff("s",Session("diaoyu"),now())
if abc<90 or abc>115 then 
	Response.Write "<script Language=Javascript>alert('你想作什麼呢？這裡不能作弊的！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 體力,狀態,操作時間 from 用戶 where id="&info(9),conn
if rs("體力")<0 or rs("狀態")="死" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Redirect "../exit.asp"
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
conn.execute "update 用戶 set 操作時間=now()  where id="&info(9)
randomize timer
r=int(rnd*6)+1
select case r
case 2
	mess="恭喜"& info(0) &"打死蝙蝠後得到一包金器，到集市賣得十萬銀子"
	conn.execute "update 用戶 set 銀兩=銀兩+100000 where id="&info(9) 
case 3
	mess="恭喜"& info(0) &"打死蝙蝠後得到一盒貴重首飾，變賣得到銀兩60000"
	conn.execute "update 用戶 set 銀兩=銀兩+60000 where id="&info(9)
case 4
	mess="恭喜"& info(0) &"打死蝙蝠後得到一串珍珠，變賣得到銀子4000兩"
	conn.execute "update 用戶 set 銀兩=銀兩+4000 where id="&info(9)
case 5
	mess="倒霉的"& info(0) &"蝙蝠沒打死，而且不小心踩到陷阱,體力減少500，內力減少200"
	conn.execute "update 用戶 set 體力=體力-500,內力=內力-200 where id="&info(9)
case 6
	mess="蝙蝠叫了幫兇,"& info(0) &"反抗遭到毒打,體力下降1000、內力下降500"
        conn.execute "update 用戶 set 體力=體力-1000,內力=內力-500 where id="&info(9)
case 7
        mess="恭喜"& info(0) &"，運氣真是好的不得了呀！打死蝙蝠後得到了一大堆黃金啊,足足值150000兩銀子！"
        conn.execute "update 用戶 set 銀兩=銀兩+150000 where id="&info(9)
case 8
        mess=""& info(0) &"運氣真是好，打死蝙蝠後居然挖出了江湖至寶靈劍水晶石！大家還不快搶！"
          conn.execute "update 用戶 set 寶物='靈劍水晶石',寶物修練=0 where id="&info(9)
        rs.close
end select

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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>消息</font>"& info(0) &"在村莊打蝙蝠："&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>打蝙蝠OK</title></head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=關閉窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>