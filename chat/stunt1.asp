<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
at1=request.form("at1")
at2=request.form("at2")
at3=request.form("at3")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if Application("jingda")=1 then
	Response.Write "<script Language=Javascript>alert('目前由於人少的原因系統被設置為禁打狀態，要想打架請找官府人員開啟！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=info(0)
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,狀態,操作時間,殺人數,體力,等級,保護 from 用戶 where id="&info(9),conn
if rs("狀態")="死" or rs("體力")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
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
	Response.Write "<script Language=Javascript>alert('連續技需要戰鬥等級5級以上才可以操作！');location.href = 'javascript:history.go(-1)';</script>"
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
rs.close
rs.open "select 會員等級,門派,體力,狀態,保護,等級 from 用戶 where 姓名='" & to1 & "'",conn
jhhy=rs("會員等級")
mp=info(5)
if rs("狀態")="死" or rs("體力")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');location.href = 'javascript:history.go(-1)';</script>"
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
if rs("等級")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('["&to1&"]為初少江湖新手，不用一這麼重的招式吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT 武功,內力 FROM 武功 WHERE 武功='" & at1 & "' or 武功='" & at2 & "' or 武功='" & at3 & "' and 門派='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你並沒有這樣的武功呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("武功")=at1 then
nei1=abs(rs("內力"))
end if
if rs("武功")=at2 then
nei2=abs(rs("內力"))
end if
if rs("武功")=at3 then
nei3=abs(rs("內力"))
end if
nei=nei1+nei2+nei3
rs.close
rs.open "select 武功,攻擊,內力 from 用戶 where id="&info(9),conn
lxjwg1=rs("武功")
lxjgj1=rs("攻擊")
if rs("內力")<nei  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('內力不夠，發不了連續技！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select 武功,攻擊,體力 from 用戶 where 姓名='" & to1 & "'",conn
lxjwg2=rs("武功")
lxjgj2=rs("攻擊")
killer=int(((lxjwg1+lxjgj1)-(lxjwg2+lxjgj2)+nei/10)/5)
'殺傷小於100
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
	conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',now(),'" & info(0) & "','連續技')"
	e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦"
	call boot(to1)
end if
stunt=info(0) & "運足內力對" & to1 & "使用了【" & at1 & "+" & at2 & "+" & at3 & "】的一連串華麗連續技，殺傷" & killer & e

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
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="對"
sd(199)="<font color=red>【連續技】"& stunt &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('恭喜，您的連續技已經發招完成！');location.href = 'javascript:history.go(-1)';</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
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
 