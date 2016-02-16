<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	response.end
end if
if Application("ljjh_gg")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子來晚了，狗狗的肉都讓人給喫了...');</script>"
	response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize timer
r=int(rnd*3)+1
if r=2 then
rs.open "select id from 物品 where 物品名='冰雪秘籍',會員 and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,會員) values ('冰雪秘籍',"&info(9)&",'物品',0,0,0,0,0,1,200000,99008,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
yb="["&info(0)&"]下手好快啊,一下子就把狗狗打死了,還從狗狗窩裡找到一頁《冰雪秘籍》<img src='../hcjs/jhjs/001/99008.gif' border='0'>大家下手要快啊!"
else
rs.open "select 等級,體力加 FROM 用戶 WHERE id="&info(9),conn
tlj=(rs("等級")*1500+5260)+rs("體力加")
conn.execute "update 用戶 set 體力=" & tlj & ",銀兩=銀兩+500000  where id="&info(9)
yb="["&info(0)&"]下手好快啊,一下子就把狗狗打死了,狗肉生喫了一干二淨,狗骨頭還賣了賺到500000兩銀子，大家下手要快啊!"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("ljjh_gg")=0
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
sd(199)="<font color=B0D0E0>【江湖消息】</font>"&yb
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
 