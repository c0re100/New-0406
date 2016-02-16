<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.B0D0E0irect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
nickname=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	response.end
end if
if Application("ljjh_bianfu")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子來晚了，怪鳥早被人打掉了...');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*12)+2
select case r
case 1
sql="select * from 物品 where 物品名='冰火銀針' and 擁有者="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,會員) values ('冰火銀針',"&info(9)&",'暗器',0,0,2000,-2000,-3000,1,20000,1300,"&aaao&")"
	else
		
		conn.execute "update 物品 set 銀兩=20000,數量=數量+1,會員="&aaao&" where 物品名='冰火銀針' and 擁有者="&info(9)
	end if
mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到一支冰火銀針，可作為暗器使用，殺傷體力3000、內力2000"
case 2
	mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到一包金器，到集市賣得十萬銀子"
	conn.execute "update 用戶 set 銀兩=銀兩+100000 where id="&info(9) 
case 3
	mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到一盒貴重首飾，變賣得到銀兩60000"
	conn.execute "update 用戶 set 銀兩=銀兩+60000 where id="&info(9)
case 4
	mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到一串珍珠，變賣得到銀子4000兩"
	conn.execute "update 用戶 set 銀兩=銀兩+4000 where id="&info(9)
case 5
	mess="倒霉的<font color=B0D0E0>【"& nickname &"】</font>怪鳥沒打死，而且不小心踩到陷阱,體力減少500，內力減少200"
	conn.execute "update 用戶 set 體力=體力-500,內力=內力-200 where id="&info(9)
case 6
	mess="怪鳥叫了幫兇,<font color=B0D0E0>【"& nickname &"】</font>反抗遭到毒打,體力下降1000、內力下降500"
        conn.execute "update 用戶 set 體力=體力-1000,內力=內力-500 where id="&info(9)
case 7
        mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>運氣真是好的不得了呀！打死怪鳥後得到了一大堆黃金啊,足足值1000000兩銀子！"
        conn.execute "update 用戶 set 銀兩=銀兩+1000000 where id="&info(9)
case 11
       mess="<font color=B0D0E0>【"& nickname &"】</font>運氣真是好，打死怪鳥後得到了<font color=B0D0E0>江湖至寶靈劍水晶石</font>！大家還不快搶！"
      conn.execute "update 用戶 set 寶物='靈劍水晶石',寶物修練=0 where id="&info(9)
case 9

sql="select id from 物品 where 物品名='回命丹' and 擁有者="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,會員) values ('回命丹',"&info(9)&",'藥品',0,0,2000,0,100000,1,200000,509,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到稀世奇珍[回命丹]，服用可增加體力100000"
case 10
sql="select id from 物品 where 物品名='天山雪蓮' and 擁有者="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,會員) values ('天山雪蓮',"&info(9)&",'藥品',0,0,3000,100000,0,1,200000,500,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&"  where id="& id &" and 擁有者="&info(9)
	end if
mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到稀世奇珍[天山雪蓮]，服用可增加內力100000"
case 11
sql="select id from 物品 where 物品名='毒蛤蟆' and 擁有者="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,堅固度,類型,攻擊,防御,內力,體力,數量,銀兩,說明,會員) values ('毒蛤蟆',"&info(9)&",'毒藥',0,0,4000,0,10000,1,200000,1200,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&"  where id="& id &" and 擁有者="&info(9)
	end if
mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到罕見毒物[毒蛤蟆]，可殺傷對方體力10000"
case 12
sql="select id from 物品 where 物品名='神龍秘籍' and 擁有者="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,堅固度,類型,攻擊,防御,內力,體力,數量,銀兩,說明,會員) values ('神龍秘籍',"&info(9)&",'物品',0,0,4000,0,10000,1,200000,99007,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&"  where id="& id &" and 擁有者="&info(9)
	end if
mess="恭喜<font color=B0D0E0>【"& nickname &"】</font>打死怪鳥後得到一本《神龍秘籍》<img src='../hcjs/jhjs/001/99007.gif' border='0'>，內有覆雨翻雲的秘招！"
case else
 conn.execute "update 用戶 set 寶物='靈劍水晶石',寶物修練=0 where id="&info(9)
 mess="<font color=B0D0E0>【"& nickname &"】</font>運氣真是好，打死怪鳥後得到了<font color=B0D0E0>江湖至寶靈劍水晶石</font>！大家還不快搶！"
       
end select

Application.Lock
Application("ljjh_bianfu")=0
Application.UnLock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【江湖消息】</font>"&mess
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
