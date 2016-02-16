<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
Response.Write "<script Language=Javascript>alert('提示：你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
response.end
end if
if Application("ljjh_bianfu")=0 then
Response.Write "<script Language=Javascript>alert('提示：你來晚了,寵物早溜了..');</script>"
response.end
end if
name=lcase(trim(request("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*8)+2
select case r
case 1
mess="<font color=red>【"& nickname &"】</font>拿起棍子想打"&name&"的寵物，結果["&name&"]的寵物獸性大發，拼命奔跑，<font color=red>【"& nickname &"】</font>追了半天都沒追上，真糗人啊,體力消耗100，內力損耗200！"
conn.execute "update 用戶 set 體力=體力-100,內力=內力-200 where id="&info(9)
case 2,3,4,6,7,8,9,10
sql="select id from 寵物狀態 where user='"&name&"'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
Response.Write "<script Language=Javascript>alert('"&name&"的寵物早死了呀！');</script>"
response.end
else
id=rs("id")
conn.execute "delete * from 寵物狀態 where id="& id &" and user='"&name&"'"
end if
rs.close
rs.open "select id from 物品 where 物品名='子彈' and 擁有者=" &info(9),conn
If Rs.Bof OR Rs.Eof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,內力,體力,數量,銀兩,會員) values ('子彈',"&info(9)&",'彈藥',0,0,100,2012,0,0,1,2000000,"&aaao&" )"
else
id=rs("id")
conn.execute "update 物品 set 銀兩=200000,數量=數量+1,會員="&aaao&" where 物品名='子彈' and 擁有者="&info(9)
end if
mess="<font color=red>【"& nickname &"】</font>拿起棍子對著["&name&"]的寵物一陣猛揍，["&name&"]的寵物終於慘叫兩聲後倒在地上，【"& nickname &"】從["&name&"]的寵物尸體裡找到一顆子彈!"
case else
mess="<font color=red>【"& nickname &"】</font>拿起棍子想打"&name&"的寵物，結果["&name&"]的寵物獸性大發，拼命奔跑，<font color=red>【"& nickname &"】</font>追了半天都沒追上，真糗人啊,體力消耗100，內力損耗200！"
conn.execute "update 用戶 set 體力=體力-100,內力=內力-200 where id="&info(9)
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
sd(196)="660099"
sd(197)="660099"
sd(198)="對"
sd(199)="<font color=red>【江湖消息】</font>"&mess
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
