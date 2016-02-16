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
if Application("ljjh_yb")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子來晚了，元寶沒有了...');</script>"
	response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 銀兩=銀兩+"& Application("ljjh_yb") &" where id="&info(9)
yb="今天真是走運，出門都有錢砸到我的腦袋上，呵。。。因此：["&name&"]得到:<img src='img/251.gif'>"&Application("ljjh_yb")&"兩!"
Application.Lock
Application("ljjh_yb")=0
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
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="對"
sd(199)="<font color=FFD060>【江湖消息】</font>"&yb
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
 