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
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"diaoyu/diao.asp")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
abc=DateDiff("s",Session("diaoyu"),now())
if abc<30 or abc>40 then 
	Response.Write "<script Language=Javascript>alert('你想作什麼呢？這裡不能作弊的！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('請等上"&ss&"分再來釣魚吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用戶 set 操作時間=now()  where id="&info(9)
randomize timer
r=int(rnd*12)+1
'r=9
select case r
case 1
	mess="恭喜"& info(0) &"釣到一條大鯊魚，可作為暗器使用，殺傷體力1200、內力800"
	rs.open "select id from 物品 where 物品名='大鯊魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大鯊魚',"&info(9)&",'暗器',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=20000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 2
	mess="恭喜"& info(0) &"釣到一條娃娃魚，到集市賣得5萬銀子"
	conn.execute "update 用戶 set 銀兩=銀兩+50000 where id="&info(9) 
case 3
	mess="恭喜"& info(0) &"釣到一隻大草魚，喫下可以增加體力300、內力50"
	rs.open "select id from 物品 where 物品名='大草魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大草魚',"&info(9)&",'藥品',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=2000,數量=數量+1,會員="&aaao&" where id="&id &" and 擁有者="&info(9)
	end if
	rs.close
case 4
	mess="恭喜"& info(0) &"釣到一條小鯉魚，喫下可以增加體力100、內力30"
	rs.open "select id from 物品 where 物品名='小鯉魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('小鯉魚',"&info(9)&",'藥品',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=500,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 5
	mess="唉！！"& info(0) &"魚沒釣到，摔到河裡體力減少300，內力減少100"
	conn.execute "update 用戶 set 體力=體力-300,內力=內力-100 where id="&info(9)
case 6
	mess= info(0) &"偷釣魚塘的魚塘被主人發現，一陣毆打，體力下降500、內力下降200"
	conn.execute "update 用戶 set 體力=體力-500,內力=內力-200 where id="&info(9)
case 7
	mess="恭喜"& info(0) &"運氣真是好的BiangBiang聲呀！釣到大鯊魚、大草魚、小鯉魚各一條！"

	rs.open "select id from 物品 where 物品名='大鯊魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大鯊魚',"&info(9)&",'暗器',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=20000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close

	rs.open "select id from 物品 where 物品名='大草魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大草魚',"&info(9)&",'藥品',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=2000,數量=數量+1,會員="&aaao&" where id="&id &" and 擁有者="&info(9)
	end if
	rs.close

	rs.open "select id from 物品 where 物品名='小鯉魚',會員="&aaao&"  and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('小鯉魚',"&info(9)&",'藥品',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=500,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 8

   mess="你在河邊撿到一本破舊的《排雲秘籍》<img src='../hcjs/jhjs/001/99004.gif' border='0'>，"&info(0)&"你真是走運啊."
rs.open "select id from 物品 where 物品名='排雲秘籍' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('排雲秘籍',"&info(9)&",'物品',0,0,0,0,0,1,200000,99004,99004,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if

rs.close
case 9
	mess="恭喜"& info(0) &"在潛水時發現了一個流星，"& info(0) &"開心極了!"
	rs.open "select id from 物品 where 物品名='流星' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('流星',"&info(9)&",'物品',0,0,1000,111111,111111,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=1000000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 10
	mess="恭喜"& info(0) &"在尋找掉進水裡的銀兩時發現了一個流星淚，真是幸運!"
	rs.open "select id from 物品 where 物品名='流星淚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('流星淚',"&info(9)&",'物品',0,0,1000,2003,2003,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=1000000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 11
	mess="恭喜"& info(0) &"在河邊發現了一個龍珠，真是驚喜!"
	rs.open "select id from 物品 where 物品名='龍珠' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('龍珠',"&info(9)&",'物品',0,0,100,50000,50000,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=1000000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 12
	mess="恭喜"& info(0) &"在釣魚的時候，有賊人送了10萬會費給他，"& info(0) &"並沒有報官，反而把那10萬會員放進自己的口袋裡。"
	conn.execute "update 用戶 set 會員費=會員費+100000 where id="&info(9) 
case 13
	mess="恭喜"& info(0) &"釣到一個盒子，內裡全都是金幣，經數算後，金幣總數竟有10萬個，"& info(0) &"這一生可以靠這些金牌過活，真是令人羨慕。"
	conn.execute "update 用戶 set 金幣=金幣+100000 where id="&info(9) 
case else
   mess="你閑著沒事往河邊跳，不料被你在河底發現一塊特種金屬<img src='../hcjs/jhjs/001/2003303.gif' border='0'>，"&info(0)&"你的賊眼好尖啊!嘿嘿!"
rs.open "select id from 物品 where 物品名='特種金屬' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('特種金屬',"&info(9)&",'物品',0,0,0,0,0,1,10000,2003303,2003303,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if

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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=#ff0000>消息</font>"& info(0) &"在河邊釣魚："&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>釣魚成功!</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif">
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