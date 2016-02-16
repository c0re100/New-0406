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
rs.open "select 體力,狀態 from 用戶 where id="&info(9),conn
if rs("體力")<0 or rs("狀態")="死" then Response.Redirect "../exit.asp"
sja=session("wabao")
randomize timer
r=int(rnd*12)+8
  if session("yuokaa")="" then
     Response.Redirect "diao.asp"
     response.end                                                           
  end if
session("wabao")=""
session("yuokaa")=""
select case r

case 1
	mess="恭喜"&info(0)&"挖到一把飛鳳劍，可作為暗器使用，殺傷體力80000、內力10000"
rs.close
rs.open "select id from 物品 where 物品名='飛鳳劍' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,說明,sm,堅固度,攻擊,防御,體力,內力,數量,會員) values ('飛鳳劍',"&info(9)&",'暗器',1000,1000,100,0,0,80000,10000,1,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
				
		        end if
case 2
	mess="恭喜"&info(0)&"挖到一支玉發簪，可以用做裝飾品，加防御1000、攻擊1000"
rs.close
rs.open "select id from 物品 where 物品名='玉發簪' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,sm,等級,會員) values ('玉發簪',"&info(9)&",'裝飾',1000,100,0,0,1000,1,200000,1500,1500,45,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
				
		        end if
case 3
	mess="恭喜"&info(0)&"挖到一支<img src='../hcjs/jhjs/001/135.gif' border=0 >千年人參，服用可增加內力100000點"
rs.close
rs.open "select id from 物品 where 物品名='千年人參' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,內力,體力,攻擊,防御,說明,sm,數量,銀兩,會員) values ('千年人參',"&info(9)&",'藥品',2000,100000,0,0,0,135,135,1,2000000,"&aaao&")"
			
                        else
id=rs("id") 
				conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
				
		        end if
case 4
	mess="恭喜"&info(0)&"挖到一隻<img src='../hcjs/jhjs/001/136.gif' border=0 >萬年靈芝，服用可增加體力200000點"
rs.close
rs.open "select id from 物品 where 物品名='萬年靈芝' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,攻擊,防御,堅固度,說明,sm,數量,銀兩,會員) values ('萬年靈芝',"&info(9)&",'藥品',0,200000,0,0,100,136,136,1,5000000,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
				
		        end if
case 5
rs.close
rs.open "select id from 物品 where 物品名='魔力鑽石' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,會員) values ('魔力鑽石',"&info(9)&",'法器',5000,0,0,0,0,1,200000,1100,1100,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
   mess="你在一塊石頭下面挖到一顆魔力鑽石，"&info(0)&"你的眼光真是尖啊."
case 6
	mess="恭喜"&info(0)&"挖到一把金鑰匙，變賣得到銀兩50000"

   conn.execute("update 用戶 set 銀兩=銀兩+50000 where id="&info(9))

case 7
	mess="恭喜"&info(0)&"挖到一盒死人骨頭，氣死了，道德下降20"
    conn.execute("update 用戶 set 道德=道德-20 where id="&info(9))

case 8
	mess="倒霉的"&info(0)&"寶藏沒找到，而且不小心踩到陷阱,內力減少200"
   conn.execute("update 用戶 set 內力=內力-200 where id="&info(9))

case 9

mess="恭喜"&info(0)&"挖到一包金器，到集市賣得十萬銀子"

    conn.execute "update 用戶 set 銀兩=銀兩+100000 where id="&info(9)
case 10
        mess=""&info(0)&"運氣真是好，居然挖出了江湖至寶靈劍水晶石！"
        conn.execute "update 用戶 set 寶物='靈劍水晶石',寶物修練=0 where id="&info(9)
case 11
rs.close
rs.open "select id from 物品 where 物品名='錯筋秘籍' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,會員) values ('錯筋秘籍',"&info(9)&",'物品',0,0,0,0,0,1,200000,99001,99001,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
   mess="你在一塊石頭下面挖到一本破舊的《錯筋秘籍》<img src='../hcjs/jhjs/001/99001.gif' border='0'>，"&info(0)&"你真是走運啊."
case 12
rs.close
rs.open "select id from 物品 where 物品名='水晶' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,會員) values ('水晶',"&info(9)&",'物品',0,0,0,0,0,1,10000,2003307,2003307,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
   mess="你在一塊石頭下面挖到一塊水晶<img src='../hcjs/jhjs/001/2003307.gif' border='0'>，"&info(0)&"你還不快取出來呀,此物可以用來煆造絕世武器呀,快收藏好！"
case else
    mess="強盜又來搶劫,"&info(0)&"反抗遭到毒打,乖乖交出1000兩"
   conn.execute("update 用戶 set 存款=存款-1000 where id="&info(9))
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
sd(194)="消息"
sd(195)="大家"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=red>【消息】</font><font color=FFCFCF>"&info(0)&"</font>在後山挖寶藏：<font color=FFCFCF>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>挖寶貝OK</title></head>

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
