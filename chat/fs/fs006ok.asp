<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得100點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<9 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=9-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
conn.execute "update 用戶 set 法力=法力-100,操作時間=now() where id="&info(9)
randomize()
'rnd1=int(rnd()*3)
r=int(rnd*7)+4
select case r
case 1
rs.close
rs.open "select id FROM 物品 WHERE 物品名='冰水' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('冰水',"&info(9)&",'藥品',0,0,100,151,151,0,150,150,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一個冰水，消耗法力100."
case 2
rs.close
rs.open "select id FROM 物品 WHERE 物品名='礦石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('礦石',"&info(9)&",'物品',0,0,100,2014,2014,0,0,0,1,800,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一個礦石，消耗法力100."
case 3
rs.close
rs.open "select id FROM 物品 WHERE 物品名='大鯊魚' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大鯊魚',"&info(9)&",'暗器',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一個大鯊魚，消耗法力100."
case 4
rs.close
rs.open "select id FROM 物品 WHERE 物品名='大草魚' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大草魚',"&info(9)&",'藥品',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一個大草魚，消耗法力100."

case 5
rs.close
rs.open "select id FROM 物品 WHERE 物品名='小鯉魚' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('小鯉魚',"&info(9)&",'藥品',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一個小鯉魚，消耗法力100."
case 6
rs.close
rs.open "select id FROM 物品 WHERE 物品名='老虎肉' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('老虎肉',"&info(9)&",'藥品',0,0,100,400,400,0,150,150,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一塊老虎肉，消耗法力100."
case 7
rs.close
rs.open "select id FROM 物品 WHERE 物品名='飛鳳劍' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('飛鳳劍',"&info(9)&",'暗器',0,0,100,1000,1000,0,10000,80000,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你施展法術<bgsound src=wav/Phant008.wav loop=1>變出了一把飛鳳劍，消耗法力100."

case else
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>你<bgsound src=wav/Phant008.wav loop=1>運氣真是差差呀，什麼都沒變出來."
end select
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>【百變術】"& banbianshu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>