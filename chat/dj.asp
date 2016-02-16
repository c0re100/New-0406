<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
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
if Application("ljjh_dj")=0 then
Response.Write "<script Language=Javascript>alert('提示：你來晚了!');</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*28)
select case r
case 1
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把青鋼刀<img src='../hcjs/jhjs/001/11001.gif' border=0 width=32 height=32  >"
rs.open "select id from 物品 where 物品名='青鋼刀' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('青鋼刀',"&info(9)&",'右手',150,100,0,0,0,1,20000,11001,5,11001,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 2
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把雁翎刀<img src='../hcjs/jhjs/001/11002.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='雁翎刀' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('雁翎刀',"&info(9)&",'右手',450,220,0,0,0,1,20000,11002,18,11002,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
case 3
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把雁翎刀<img src='../hcjs/jhjs/001/11003.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='龍鱗刀' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('龍鱗刀',"&info(9)&",'右手',650,350,0,0,0,1,20000,11003,28,11003,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 4
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把雁翎刀<img src='../hcjs/jhjs/001/11004.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='漏影刀' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('漏影刀',"&info(9)&",'右手',860,420,0,0,0,1,20000,11004,38,11004,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 5
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把萬人斬<img src='../hcjs/jhjs/001/11005.gif' border=0 width=32 height=32  >"
rs.open "select id from 物品 where 物品名='萬人斬' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('萬人斬',"&info(9)&",'右手',4000,800,0,0,0,1,20000,11005,60,11005,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
case 6
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個流星<img src='../hcjs/jhjs/001/111111.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='流星' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('流星',"&info(9)&",'物品',0,1000,0,0,0,1,10000000,111111,0,111111,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 7
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把龍淵劍<img src='../hcjs/jhjs/001/11007.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='龍淵劍' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('龍淵劍',"&info(9)&",'右手',35000,3190,0,0,0,1,20000,11007,145,11007,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 8
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把青光劍<img src='../hcjs/jhjs/001/11008.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='青光劍' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('青光劍',"&info(9)&",'右手',38000,4070,0,0,0,1,20000,11008,185,11008,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
case 9
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一把無雙劍<img src='../hcjs/jhjs/001/11006.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='無雙劍' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('無雙劍',"&info(9)&",'右手',25000,1870,0,0,0,1,20000,11006,85,11006,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 10
mess=info(0)&"打死怪物獲得20點經驗!"
conn.execute "update 用戶 set allvalue=allvalue+20,體力=體力-100 where id="&info(9)
case 11
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個珍珠戒指<img src='../hcjs/jhjs/001/6100.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='珍珠戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('珍珠戒指',"&info(9)&",'裝飾',0,80,0,0,120,1,20000,6100,12,6100,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 12
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個藍星戒指<img src='../hcjs/jhjs/001/6101.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='藍星戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('藍星戒指',"&info(9)&",'裝飾',0,100,0,0,150,1,20000,6101,20,6101,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 13
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個紅星戒指<img src='../hcjs/jhjs/001/6102.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='紅星戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('紅星戒指',"&info(9)&",'裝飾',0,180,0,0,400,1,20000,6102,40,6102,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 14
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個血液戒指<img src='../hcjs/jhjs/001/6103.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='血液戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('血液戒指',"&info(9)&",'裝飾',0,280,0,0,600,1,20000,6103,50,6103,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 15
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個月光戒指<img src='../hcjs/jhjs/001/6104.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='月光戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('月光戒指',"&info(9)&",'裝飾',0,380,0,0,800,1,20000,6104,60,6104,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 16
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個太陽戒指<img src='../hcjs/jhjs/001/6105.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='太陽戒指' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('太陽戒指',"&info(9)&",'裝飾',0,480,0,0,1000,1,20000,6105,80,6105,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 17
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個珍珠項鏈<img src='../hcjs/jhjs/001/6106.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='珍珠項鏈' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('珍珠項鏈',"&info(9)&",'裝飾',0,580,0,0,1400,1,20000,6106,100,6106,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 18
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個意念項鏈<img src='../hcjs/jhjs/001/6107.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='意念項鏈' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('意念項鏈',"&info(9)&",'裝飾',0,880,0,0,1800,1,20000,6107,120,6107,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 19
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個永恆項鏈<img src='../hcjs/jhjs/001/6108.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='永恆項鏈' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('永恆項鏈',"&info(9)&",'裝飾',0,1480,0,0,2800,1,20000,6108,160,6108,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 20
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個青銅盾牌<img src='../hcjs/jhjs/001/2100.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='青銅盾牌' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('青銅盾牌',"&info(9)&",'左手',0,500,0,0,1000,1,20000,2100,50,2100,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
rs.close
case 21
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個明光盾牌<img src='../hcjs/jhjs/001/2101.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='明光盾牌' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('明光盾牌',"&info(9)&",'左手',0,1500,0,0,3000,1,20000,2101,80,2101,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 22
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個水晶盾牌<img src='../hcjs/jhjs/001/2102.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='水晶盾牌' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('水晶盾牌',"&info(9)&",'左手',0,2500,0,0,4000,1,20000,2102,140,2102,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 23
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個生化盾牌<img src='../hcjs/jhjs/001/2103.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='生化盾牌' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('生化盾牌',"&info(9)&",'左手',0,4500,0,0,6000,1,20000,2103,180,2103,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 24
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個極藥<img src='../hcjs/jhjs/001/9100.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='極藥' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('極藥',"&info(9)&",'藥品',0,100,0,10000000,0,1,20000,9100,0,9100,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
case 25
randomize()
rr=int(rnd*10000)
if rr=5 then
mess="恭喜<font color=B0D0E0>"&info(0)&"</font>打死怪物得到一個龍珠<img src='../hcjs/jhjs/001/50000.gif' border=0  width=32 height=32 >"
rs.open "select id from 物品 where 物品名='龍珠' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,sm,會員) values ('龍珠',"&info(9)&",'物品',0,100,0,0,0,1,10000000,50000,0,50000,"&aaao&")"
else
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

end if
rs.close
else

mess=info(0)&"打死怪物得到一把金子，共計<img src='img/251.gif'>3萬兩!"
conn.execute "update 用戶 set 銀兩=銀兩+30000,體力=體力-100 where id="&info(9)
end if
case else
mess=info(0)&"打死怪物得到一把金子，共計<img src='img/251.gif'>3萬兩!"
conn.execute "update 用戶 set 銀兩=銀兩+30000,體力=體力-100 where id="&info(9)

end select

Application.Lock
Application("ljjh_dj")=0
Application.UnLock
conn.close
set conn=nothing
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
%>
