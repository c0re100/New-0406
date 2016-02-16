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
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if Application("ljjh_qiang")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你來晚了，手槍被人搶走了...');</script>"
	response.end
end if
Application.Lock
Application("ljjh_qiang")=0
Application.UnLock
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 物品 where 物品名='精制手槍' and 擁有者=" &info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,sm,會員) values ('精制手槍',"&info(9)&",'槍械',0,0,1000,0,0,1,2000000,1400,1400,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+1,會員="&aaao&" where 物品名='精制手槍' and 擁有者="&info(9)
	end if

qiang="["&name&"]真是厲害，居然從眾高手中搶先得到一把精制手槍,可惜沒有子彈，要子彈去打土匪吧或許能找到一些!"
rs.close
	set rs=nothing
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<font color=9FDF70>【江湖消息】</font>"&qiang
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
