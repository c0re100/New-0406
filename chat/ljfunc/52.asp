<%'教訓孩子
function hzb(to1,toname)
if info(11)="有過" then
	hzb=info(0)&",有沒搞錯呀，你的孩子就偷東西，你還教訓人家！"
	exit function
end if 
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('教訓孩子對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 等級,小孩 FROM 用戶 WHERE  id="&info(9),conn
fromdj=rs("等級")
if rs("小孩")="有過" then
	hzb=info(0)&",有沒搞錯呀，你的孩子就偷東西，你還教訓人家！"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
rs.close
rs.open "SELECT id,等級,小孩,孩名 FROM 用戶 WHERE 姓名='"&toname&"'",conn
idd=rs("id")
to1dj=rs("等級")
to1dd=rs("孩名")
if rs("小孩")="無" then
	hzb=info(0)&",你發什麼神經,"&toname&"還沒孩子呢！"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
if rs("小孩")="有" then
	hzb=info(0)&","&toname&"的孩子["&to1dd&"]沒偷過東西,你不能亂訓!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 

if rs("小孩")="有過" then
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update 用戶 set 小孩='有' where 姓名='"&toname&"'"
	conn.execute "update 用戶 set 道德=道德+100  where id="&info(9)
	hzb="哈哈！"&toname&"的孩子["&to1dd&"]被"&info(0)&"教訓了一頓，丟人現眼啦!("&info(0)&"道德提升100點)"

Case 2
	if todj>fromdj then
		conn.execute "update 用戶 set 魅力=魅力-100  where id="&info(9)
		conn.execute "update 用戶 set 小孩='有' where 姓名='"&toname&"'"
		hzb="嘿嘿！"&toname&"的孩子["&to1dd&"]異常敏捷,"&info(0)&"根本不是孩子的對手,終於讓小偷"&toname&"給跑了("&info(0)&"魅力下降100點)"
	else
		hzb=info(0)&"正想逮住"&toname&"的孩子["&to1dd&"],仔細一瞧,原來是大姪子,於是把"&toname&"的乖孩子放了!"
	end if 
Case 3
	conn.execute "update 用戶 set 魅力=魅力+100,道德=道德+100,銀兩=銀兩+2000000,內力=內力-200  where id="&info(9)
	conn.execute "update 用戶 set 狀態='獄',登錄=now()+3,小孩='有',內力=內力/2,體力=體力/2 where 姓名='"&toname&"'"
	call boot(toname)
	hzb="哇！,"&info(0)&"將"&toname&"的孩子["&to1dd&"]和大人一起大罵了一頓！"&toname&"覺得沒面子到家了，躲到廁所不敢出來了。("&info(0)&"獲得獎金2000000兩,道德提升100點,魅力提升100點,內力下降200點。"&toname&"體力和內力下降一半)"

Case 4
	conn.execute "update 用戶 set 小孩='有' where 姓名='"&toname&"'"
	conn.execute "update 用戶 set 魅力=魅力-100  where id="&info(9)
rs.close
	rs.open "select id,物品名,類型,內力,體力,攻擊,防御,堅固度,等級,銀兩,說明 from 物品 where  數量>0 and 裝備=false and  類型<>'卡片' and 擁有者="&idd ,conn
	id=rs("id")
	wupin=rs("物品名")
	lx=rs("類型")
	nl=rs("內力")
	tl=rs("體力")
	gj=rs("攻擊")
	fy=rs("防御")
	yin=rs("銀兩")
        jgd=rs("堅固度")
	dj=rs("等級")
	if lx="鮮花" or lx="右手" or lx="左手" or lx="物品" or lx="法器" or lx="法寶" or lx="槍械" or lx="彈藥" or lx="藥品" or lx="雙腳" or lx="頭盔" or lx="裝飾" or lx="盔甲" or lx="暗器" then
		sm=rs("說明")
	else
		sm=rs("類型")
	end if
	rs.close
    conn.execute "update 物品 set 數量=數量-1 where id="&id
	rs.open "select id from 物品  where 裝備=false and 物品名='" & wupin & "' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品 (物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,等級) values ('"&wupin&"','"&info(0)&"','"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"& sm &"',"& dj &")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=" & yin & ",數量=數量+1,說明='"& sm &"' where id="&id
	end if
	hzb="嘻！"&toname&"的孩子["&to1dd&"]非常靈活,躲避了"&info(0)&"，倉皇而逃,身上的物品["&wupin&"]掉了下來,被人家撿到(魅力下降100點)"

case else
	hzb=info(0)&"正想教訓"&toname&"的孩子["&to1dd&"],仔細一瞧,原來是大姪子,於是把"&toname&"的孩子給放了!"
End Select

end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>