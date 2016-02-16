<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
at=request.form("at")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=info(0)
compact=""
compact1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間,狀態,殺人數,等級,保護,二婚 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('警告：請等"& s &"秒再發招,可別累著！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("狀態")="死" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("殺人數")>=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('壞事作盡，殺人數滿，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("等級")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('夫妻合體技需要戰鬥等級5級以上才可以操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('請取消練功保護再操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
peiou=rs("二婚")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(PEIOU)&" ")=0 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的二婚["&peiou&"]沒有在聊天室中不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select 門派,狀態,保護,會員等級,等級 from 用戶 where 姓名='" & to1 & "'",conn
if rs("狀態")="死" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('對方已經死了，真死？假死？！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]再在練功保護中，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
jhhy=rs("會員等級")
ntnt=rs("等級")
mp=info(5)
if rs("等級")<=2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]為初少江湖新手，不用一這麼重的招式吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM 合體技 WHERE 姓名男='" & info(0) & "' or 姓名女='" & info(0) & "' and 合體技='" & at & " '",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你並沒有["&at&"]這樣的夫妻合體技呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei=abs(rs("內力"))
rs.close
rs.open "select 武功,攻擊 from 用戶 where 姓名='" & peiou & "'",conn
	htwg1=rs("武功")
	htgj1=rs("攻擊")
rs.close	
rs.open "select 內力,武功,攻擊 from 用戶 where id="&info(9),conn
if rs("內力")<nei then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('內力不夠，發不了這招合何技！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

htwg2=rs("武功")
htgj2=rs("攻擊")
rs.close
rs.open "select 武功,防御,體力 from 用戶 where 姓名='" & to1 & "'",conn
towg=rs("武功")
tofy=rs("防御")
killer=int((((htwg1+htwg2+htgj1+htgj2)-towg-tofy)+nei/10)/7)
'如果殺傷力小到100隨機數
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update 用戶 set 內力=內力-" & nei & ",操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 體力=體力-"  & killer & " where 姓名='" & to1 & "'"
e=""
if rs("體力")<-100 then
	conn.execute "update 用戶 set 殺人數=殺人數+1 where id="&info(9)
	conn.execute "update 用戶 set 狀態='死' where 姓名='" & to1 & "'"
	conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',now(),'" & from1 & "','合體技')"
	call boot(to1)
	e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦"
end if
compact="["&from1 & "]運足" & nei & "點內力與{" & peiou & "}一起對(" & to1 & ")使用了名為" & at & "的夫妻合體技，殺傷" & killer & e
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
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【合體技】"& compact &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('恭喜，您的夫妻合體技已經完成！');location.href = 'javascript:history.go(-1)';</script>"
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
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%> 