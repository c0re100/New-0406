<%'抓小偷
function pk(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('抓小偷對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以抓小偷！');}</script>"
	Response.End
end if
if left(info(8),2)="小偷" then
	pk=info(0)&",有沒搞錯呀，你就是小偷，還抓什麼小偷呀！"
	exit function
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 等級 FROM 用戶 WHERE  id="&info(9),conn
fromdj=rs("等級")
rs.close
rs.open "SELECT 等級,門派,職業 FROM 用戶 WHERE  姓名='"&toname&"'",conn
to1dj=rs("等級")
if rs("門派")="官府" then
   pk=info(0)&",你有沒有搞錯呀，官府哪會有小偷!"
   rs.close
   set rs=nothing
   conn.close
   set conn=nothing
   exit function
end if

if rs("職業")<>"小偷(在逃)" then
	pk=info(0)&",你發什麼神經,"&toname&"可沒偷過東西,你抓他做什麼!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 

if rs("職業")="小偷(在逃)" then
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update 用戶 set 職業='戰士' where 姓名='"&toname&"'"
	conn.execute "update 用戶 set 道德=道德+200  where id="&info(9)
	pk="小偷"&toname&"在"&info(0)&"的教導下,決定投案自首,改過自新不再做小偷!("&info(0)&"道德提升200點)"

Case 2
	if todj>fromdj then
		conn.execute "update 用戶 set 魅力=魅力-100  where id="&info(9)
		conn.execute "update 用戶 set 職業='戰士' where 姓名='"&toname&"'"
		pk="小偷"&toname&"武功高強,"&info(0)&"根本不是小偷的對手,終於讓小偷"&toname&"給跑了("&info(0)&"魅力下降100點)"
	else
		pk=info(0)&"正想逮住小偷"&toname&",仔細一瞧,原來是自家親戚,於是把小偷"&toname&"放了!"
	end if 
Case 3
	conn.execute "update 用戶 set 魅力=魅力+100,道德=道德+100,銀兩=銀兩+10000,內力=內力-200  where id="&info(9)
	conn.execute "update 用戶 set 職業='戰士',狀態='死' where 姓名='"&toname&"'"
	call boot(toname)
	pk="經過一番苦戰,捕快"&info(0)&"終於將小偷"&toname&"逮捕歸案!(捕快獲得獎金10000兩,道德提升100點,魅力提升100點,內力下降200點)"

Case 4
	conn.execute "update 用戶 set 職業='戰士' where 姓名='"&toname&"'"
	conn.execute "update 用戶 set 魅力=魅力-100  where id="&info(9)
	rs.close
rs.open "select id from 用戶 where  姓名='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id,物品名,類型,內力,體力,攻擊,防御,銀兩,等級,堅固度,說明 from 物品 where  數量>0 and 裝備=false and  類型<>'卡片' and 擁有者="&idd ,conn
	id=rs("id")
	wupin=rs("物品名")
	lx=rs("類型")
	nl=rs("內力")
	tl=rs("體力")
	gj=rs("攻擊")
	fy=rs("防御")
	yin=rs("銀兩")
	dj=rs("等級")
	jgd=rs("堅固度")
	if lx="鮮花" or lx="右手" or lx="左手" or lx="物品" or lx="法器" or lx="法寶" or lx="槍械" or lx="彈藥" or lx="藥品" or lx="雙腳" or lx="頭盔" or lx="裝飾" or lx="盔甲" or lx="暗器" or lx="毒藥" then
	sm=rs("說明")
else
	sm=rs("類型")
end if
	
    conn.execute "update 物品 set 數量=數量-1,會員="&aaao&" where id="&id
rs.close
	rs.open "select id from 物品  where 裝備=false and 物品名='" & wupin & "' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品 (物品名,擁有者,類型,攻擊,防御,內力,體力,數量,銀兩,說明,等級,堅固度,會員) values ('"&wupin&"','"&info(0)&"','"&lx&"',"&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"& sm &"',"&dj&","&jgd&","&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
	end if
	pk="小偷"&toname&"非常靈活,躲避了捕快"&info(0)&"的追捕,倉皇而逃,身上的物品["&wupin&"]掉了下來,被捕快撿到(捕快魅力下降100點)"

case else
	pk=info(0)&"正想逮住小偷"&toname&",仔細一瞧,原來是自家親戚,於是把小偷"&toname&"放了!"
End Select

end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>