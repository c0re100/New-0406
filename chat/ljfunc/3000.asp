<%
'尋找法器
function aaaaaa(fn1,to1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-1000,操作時間=now() where id="&info(9)
randomize()
r=int(rnd*25)+1
select case r
case 1
rs.close
rs.open "select id FROM 物品 WHERE 物品名='狼牙棒' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('狼牙棒',"&info(9)&",'法器',0,0,1000,2001,2001,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你在"&towho&"家裡的床底下偷到了<font color=red>法器狼牙棒</font>，隨後鑽地洞跑了."
case 2
rs.close
rs.open "select id FROM 物品 WHERE 物品名='破天錐' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('破天錐',"&info(9)&",'法器',0,0,1000,2002,2002,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0)& "你闖進"&towho&"家裡把<font color=red>法器破天錐</font>從其手中搶走，真是世事不公啊."
case 3
rs.close
rs.open "select id FROM 物品 WHERE 物品名='血滴子' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('血滴子',"&info(9)&",'法器',0,0,1000,2004,2004,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你在江湖裡的一個破廟中發現了被別人丟棄的<font color=red>法器血滴子</font>."
case 4
rs.close
rs.open "select * FROM 物品 WHERE 物品名='搶劫令' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('搶劫令',"&info(9)&",'法器',0,0,1000,2005,2005,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你真是幸運，難得一見的<font color=red>法器搶劫令</font>都能被你找，真是有福之人啊."

case 5
rs.close
rs.open "select id FROM 物品 WHERE 物品名='紅寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('紅寶石',"&info(9)&",'法器',0,0,1000,2006,2006,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你發現"&towho&"口袋裡紅光閃閃,順手一摸原來是一顆<font color=red>紅寶石</font>,於是撿了塊碎石頭塞進"&towho&"的口袋,把紅寶石給盜走了."
case 6
rs.close
rs.open "select id FROM 物品 WHERE 物品名='綠寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('綠寶石',"&info(9)&",'法器',0,0,1000,2007,2007,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你發現"&towho&"口袋裡綠光閃閃,順手一摸原來是一顆<font color=red>綠寶石</font>,於是撿了塊碎石頭塞進"&towho&"的口袋,把綠寶石給盜走了."
case 7
rs.close
rs.open "select id FROM 物品 WHERE 物品名='藍寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('藍寶石',"&info(9)&",'法器',0,0,1000,2008,2008,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你發現"&towho&"口袋裡藍光閃閃,順手一摸原來是一顆<font color=red>藍寶石</font>,於是撿了塊碎石頭塞進"&towho&"的口袋,把藍寶石給盜走了."

case 8
rs.close
rs.open "select id FROM 物品 WHERE 物品名='魔力鑽石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('魔力鑽石',"&info(9)&",'法器',0,0,100,1100,1100,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你在江湖裡的一棵千年古樹的樹枝上發現一顆<font color=red>魔力鑽石</font>，"&info(0)&"你的眼光真是尖啊."
case 9
rs.close
rs.open "select id FROM 用戶 WHERE 姓名='"&toname&"'"
idd=rs("id")
rs.close
rs.open "select id FROM 物品 WHERE 物品名='生日蛋糕' and 擁有者=" & idd
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('生日蛋糕',"&idd&",'法器',0,0,100,2009,2009,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id 
	end if
aaaaaa=info(0) & "你在去"&towho&"家過生日,送了一個蛋糕給"&towho&"!"
case 10
rs.close
rs.open "select id FROM 物品 WHERE 物品名='絕情刀' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('絕情刀',"&info(9)&",'法器',0,0,1000,2010,2010,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
aaaaaa=info(0) & "你去"&towho&"家聊天,在門後發現<font color=red>一把絕情刀</font>."
case else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>