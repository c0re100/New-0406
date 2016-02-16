<%'孩子偷竊
function hza(to1,toname)
if info(11)="無" or info(13)="無" then
	hza=info(0)&"你還沒有孩子呢！"
	exit function
end if 
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('孩子偷竊對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以孩子偷竊！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 等級,小孩,操作時間 FROM 用戶 WHERE id="&info(9),conn
dj=rs("等級")
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("小孩")="無" then
	hza=info(0)&"沒孩子怎麼偷東西？"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if 
conn.execute "update 用戶 set 道德=道德-10,銀兩=銀兩-5000,小孩='有過',操作時間=now() WHERE  id="&info(9)

lxlx="("
if dj>=5 then 
	lxlx=lxlx&"類型='藥品'"
end if 
if dj>=6 then 
	lxlx=lxlx&" or 類型='毒藥'"
end if 
if dj>=7 then
	lxlx=lxlx&" or 類型='鮮花'"
end if 
if dj>=8 then
	lxlx=lxlx&" or 類型='暗器'"
end if 
if dj>=9 then 
	lxlx=lxlx&" or 類型='裝飾'"
end if 
if dj>=10 then 
	lxlx=lxlx&" or 類型='左手' or 類型='右手'"
end if 
if dj>=11 then 
	lxlx=lxlx&" or 類型='頭盔' or 類型='盔甲' or 類型='雙腳' or lx='物品' or lx='法器' or lx='法寶' or lx='槍械' or lx='彈藥'"
end if 
lxlx=lxlx&")"
rs.close
rs.open "select id from 用戶 where 姓名='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id,物品名,類型,內力,體力,攻擊,防御,銀兩,等級,堅固度,說明 from 物品 where "&lxlx&" and 數量>0  and 類型<>'卡片' and 裝備=false and 擁有者="&idd,conn
if rs.eof or rs.bof then
            hza=info(0)&"，你可真是倒霉呀，"&toname&"身上什麼也沒有，你的孩子怎麼偷？"
            rs.close
            set rs=nothing
            conn.close
            set conn=nothing
            exit function
else
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
	conn.execute "update 物品 set 數量=數量-1 where id="&id
	rs.close
	rs.open "select 物品名,id from 物品  where 裝備=false and 物品名='" & wupin & "' and 擁有者="& info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品 (物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,等級) values ('"&wupin&"',"&info(9)&",'"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&",1,"&yin&",'"&sm&"',"&dj&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,說明='"&sm&"' where id="&id
	end if
	hza=info(0)&"的孩子["&info(13)&"]<bgsound src=wav/xiaotou.wav loop=1>從"&toname&"的身上竊取["&wupin&"]1個，瞧他樂的那個樣~~~<font color=red>(降道德10點，銀兩下降5000)</font>"
end if     
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>