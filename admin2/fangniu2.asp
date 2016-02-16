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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間,等級,體力加,內力加,體力,內力,知質 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'fangniu.asp';}</script>"
	Response.End
end if
tlj=(rs("等級")*1500+5260)+rs("體力加")
nlj=(rs("等級")*640+2000)+rs("內力加")
tla=rs("體力")
nla=rs("內力")
mymoney=rs("知質")
dj=rs("等級")
if mymoney<200 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('錯誤！你目前的資質低於200,師傅怕你放牛的時候不小心掉下山崖，所以給其他人機會了！');location.href = 'fangniu.asp';}</script>"

response.end
end if
if dj<2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('錯誤！你目前的等級低於3,師傅怕你放牛的時候遇到野獸，所以不給你去了！');location.href = 'fangniu.asp';}</script>"
response.end
end if
randomize timer
r=int(rnd*10)+1
select case r
case 1
randomize timer
s=int(rnd*50)
mess="恭喜你放牛中，找到了一本武林奇書和資質都增加"& s &""
conn.execute "update 用戶 set 銀兩=銀兩+'"& s &"',知質=知質+'"& s &"',操作時間=now()  where id="&info(9)
case 2
randomize timer
s=int(rnd*50000)
	mess="恭喜你放牛中，在路邊撿到"& s &"兩"
conn.execute "update 用戶 set 銀兩=銀兩+'"& s &"',操作時間=now()  where id="&info(9)
case 3
	mess="你放牛不專心，牛都丟了，師傅罰你5000兩，魅力減少100點"
conn.execute "update 用戶 set 銀兩=銀兩-5000,魅力=魅力-100,操作時間=now()  where id="&info(9)
case 4
	mess="恭喜你放牛中，在路邊撿到一個豆子，點數增加2點"
conn.execute "update 用戶 set 銀兩=銀兩-5000,魅力=魅力-100,泡豆點數=泡豆點數+2,操作時間=now()  where id="&info(9)
case 5
randomize timer
s=int(rnd*5000)
	mess="恭喜您!經過一天的放牛，師傅給了你一本武學秘笈，使你獲得"& s &"體力,內力提升"& s &"，辛苦沒有白費的阿！！"
if tla>=tlj or nla>=nlj then
conn.execute "update 用戶 set 體力='"&tlj&"',內力='"& nlj &"',操作時間=now()  where id="&info(9)
else
conn.execute "update 用戶 set 體力=體力+'"& s &"',內力=內力+'"& s &"',操作時間=now()  where id="&info(9)
end if
case 6
rs.close
rs.open "select id from 物品 where 物品名='牛屎' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,sm,會員) values ('牛屎','"&info(9)&"','毒藥',0,0,0,0,10000,1,20000,100013,100013,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
 end if
mess="恭喜你放牛中，在路邊撿到一個牛屎,可殺傷對方體力10000"
case 7
	mess="恭喜"&info(0)&"放牛時跌到山洞裡，機緣巧合被你發現一隻萬年靈芝，服用可增加體力200000點"
rs.close
rs.open "select id from 物品 where 物品名='萬年靈芝' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,攻擊,防御,堅固度,說明,sm,數量,銀兩,會員) values ('萬年靈芝',"&info(9)&",'藥品',0,200000,0,0,100,136,136,1,5000000,"&aaao&")"
                        else 
id=rs("id")
				sql="update 物品 set 數量=數量+1,會員="&aaao&" where 物品名='萬年靈芝' and id="&id
				conn.execute(sql)
		        end if
case 8
	mess="恭喜"&info(0)&"放牛時在溪邊撿到一本《化骨掌譜》"
rs.close
rs.open "select id from 物品 where 物品名='化骨掌譜' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,攻擊,防御,堅固度,說明,sm,數量,銀兩,會員) values ('化骨掌譜',"&info(9)&",'物品',0,0,0,0,100,99002,99002,1,5000000,"&aaao&")"
                        else 
id=rs("id")
				sql="update 物品 set 數量=數量+1,會員="&aaao&" where 物品名='化骨掌譜' and id="&id
				conn.execute(sql)
		        end if
case 9
	mess="恭喜"&info(0)&"放牛時在溪邊撿到一塊合金，此物可用來煆造絕世武器!"
rs.close
rs.open "select id from 物品 where 物品名='合金塊' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,攻擊,防御,堅固度,說明,sm,數量,銀兩,會員) values ('合金塊',"&info(9)&",'物品',0,0,0,0,0,2003302,2003302,1,10000,"&aaao&")"
                        else 
id=rs("id")
				sql="update 物品 set 數量=數量+1,會員="&aaao&" where 物品名='合金塊' and id="&id
				conn.execute(sql)
		        end if
case else
rs.close
rs.open "select id from 物品 where 物品名='牛屎' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,sm,會員) values ('牛屎','"&info(9)&"','毒藥',0,0,0,0,10000,1,20000,100013,100013,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
 end if
mess="恭喜你放牛中，在路邊撿到一個牛屎,可殺傷對方體力10000"
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'fangniu.asp';}</script>"

response.end%>