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
if Application("ljjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是沒事找事呀，哪裡有妖怪？！');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 體力=體力-"&Application("ljjh_kl")&" where id="&info(9)
rs.open "select 體力 FROM 用戶 WHERE id="&info(9),conn
if rs("體力")<-100 then
	conn.execute "update 用戶 set 狀態='死',lastkick='妖怪' where id="&info(9)
        
	call boot(info(0))
	conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & info(0) & "',now(),'妖怪','因勇打妖怪所犧牲')"
	kl="<img src='pic/kl.gif'>太不幸了，["&info(0)&"]為了保護大家的安全，舍身取義，被妖怪咬的七零八落，死翹翹了   "
else
	conn.execute "update 用戶 set 銀兩=銀兩+"& Application("ljjh_kl")*55 &" where 姓名='" & info(0) & "'"
	kl="我call,英雄，真是英雄，["&info(0)&"]與妖怪大打出手  妖怪終於倒下了，而他自己也傷的不輕，官府獎勵<img src='img/251.gif'>"&Application("ljjh_kl")*55&"兩！"
	Application.Lock
	Application("ljjh_kl")=0
	Application.UnLock
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
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
sd(199)="<font color=9FDF70>【江湖消息】</font>"&kl
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
'踢人
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
end sub
%>
 