<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then
Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
at=request.form("at")
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
from1=info(0)
pet=""
pet1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,狀態,操作時間,殺人數,等級,保護 from 用戶 where id="&info(9),conn

if rs("狀態")="死" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
s=60-sj
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
if rs("等級")<=3 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('寵物攻擊需要戰鬥等級3級以上才可以操作！');location.href = 'javascript:history.go(-1)';</script>"
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
rs.open "select 會員等級,攻擊,防御,等級,體力,體力加,門派,保護 from 用戶 where 姓名='" & to1 & "'",conn
jhhy=rs("會員等級")
mp=info(5)
togj=rs("攻擊")
tofy=rs("防御")
dengji=rs("等級")
totl=rs("體力")
tlj=(dengji*1500+5260)+rs("體力加")
if rs("保護")=true then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('對方在練功保護中，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("等級")<=2 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('對方為江湖新手，還是不用一這麼重的招式吧！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
rs.close
rs.open "SELECT * FROM 寵物狀態 WHERE user='" & from1 & "'",conn
if rs.eof or rs.bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你並沒有寵物，怎麼攻擊呀！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
adde=rs("類型id")
att=abs(rs("攻擊"))
lx=rs("寵物類型")
say=rs("說明")
jn=rs("技能")
petname=rs("名字")
ap=abs(rs("經驗"))
zou=abs(rs("忠誠"))
fangyu=abs(rs("防御"))
Select Case at
case "普通攻擊"
kills=att*(ap+zou)
randomize timer
xqz=int(rnd*100)+10
xx=int(rnd*10)
kills=int(kills-tofy*togj)
if jn="體力型" then
conn.execute "update 用戶 set 體力=體力+"&xqz&" where id="&info(9)
conn.execute "update 用戶 set 體力=體力-"&xqz&" where 姓名='"&to1&"'"
aa="吸取<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>點體力"
end if
if jn="內力型" then
conn.execute "update 用戶 set 內力=內力+"&xqz&" where id="&info(9)
conn.execute "update 用戶 set 內力=內力-"&xqz&" where 姓名='"&to1&"'"
aa="吸取<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>點內力"
end if
if jn="防御型" then
conn.execute "update 用戶 set 防御=防御+"&xqz&" where id="&info(9)
conn.execute "update 用戶 set 防御=防御-"&xqz&" where 姓名='"&to1&"'"
aa="打掉<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>點防御力"
end if
if jn="武功型" then
conn.execute "update 用戶 set 武功=武功+"&xqz&" where id="&info(9)
conn.execute "update 用戶 set 武功=武功-"&xqz&" where 姓名='"&to1&"'"
aa="學會<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>點武功"
end if
if jn="法力型" then
conn.execute "update 用戶 set 法力=法力+"&xqz&" where id="&info(9)
conn.execute "update 用戶 set 法力=法力-"&xqz&" where 姓名='"&to1&"'"
aa="吸取<font color=9FDF70>"&to1&"</font><font color=red>"&xqz&"</font>點法力"
end if


if kills>10 and kills<tlj then
conn.execute "update 寵物狀態 set 體力=體力-" & int(togj/fangyu) & ",經驗=經驗+10 where user='" & from1 & "'"
conn.execute "update 用戶 set 操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 體力=體力-"  & kills & " where 姓名='" & to1 & "'"
rs.close
rs.open "select * from 用戶 where 姓名='" & to1 & "'",conn
if rs("體力")<-100 then
conn.execute "update 用戶 set 殺人數=殺人數+1 where id="&info(9)
conn.execute "update 用戶 set 狀態='死' where 姓名='" & to1 & "'"
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',now(),'" & info(0) & "','寵物攻擊')"
e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦"
call boot(to1)
end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出“普通攻擊”攻擊" & to1 & "，寵物被反擊，體力降低" & int(togj/fangyu) & "點，咬傷" & to1 & "體力" & kills & e
elseif kills>=tlj then
conn.execute "update 寵物狀態 set 體力=體力-" & int(togj/fangyu) & ",經驗=經驗+10 where user='" & from1 & "'"
conn.execute "update 用戶 set 操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 體力=體力/2,攻擊=攻擊-1000,防御=防御-1000 where 姓名='" & to1 & "'"
rs.close
rs.open "SELECT 體力 FROM 寵物狀態 WHERE user='" & from1 & "'",conn
if rs("體力")>0 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出--[普通攻擊]--攻擊<font color=9FDF70>" & to1 & "</font>，寵物被反擊，體力降低<font color=red>" & int(togj/fangyu) & "</font>點，" & to1 & "攻擊、防御各減少1000點，體力減少" & totl/2& "點"
else
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>的寵物<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]經受不起<font color=9FDF70>" & to1 & "</font>的一頓毒打，終於慢慢慢慢的倒了下去  從此江湖又少了一隻寵物！"
conn.execute "Delete * From 寵物狀態 Where user='" & from1 & "'"
'conn.execute "Delete * From 技能列表 Where 擁有者='" & from1 & "'"
conn.execute "Delete * From 道具列表 Where 擁有者='" & from1 & "'"
end if
else
if xx=9 or xx=7 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出--[普通攻擊]--攻擊<font color=red>" & to1 & "</font>。<font color=red>" & to1 & "</font>也使出渾身解數，戰鬥進行的異常激烈，最後雙方打成平手，互無損失！"
elseif xx=8 then
sql="update 寵物狀態 set 體力=體力-"& xqz &",防御=防御-"& xqz*2 &" where user='" & from1 & "'"
conn.execute sql
rs.close
rs.open "SELECT 體力 FROM 寵物狀態 WHERE user='" & from1 & "'",conn
if rs("體力")<0 then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>的寵物<img src='../myhome/sheep/image/"&say&".gif'width=49 height=46>[" & lx & "]經受不起<font color=9FDF70>" & to1 & "</font>的一頓毒打，終於慢慢慢慢的倒了下去  從此江湖又少了一隻寵物！"
conn.execute "Delete * From 寵物狀態 Where user='" & from1 & "'"
'conn.execute "Delete * From 技能列表 Where 擁有者='" & from1 & "'"
conn.execute "Delete * From 道具列表 Where 擁有者='" & from1 & "'"
else
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出--[普通攻擊]--攻擊<font color=red>" & to1 & "</font>。<font color=red>" & to1 & "</font>大喊：誰家的賴皮寵物也敢攻擊我？嘗嘗我的“痛打落水狗”絕招，" & petname & "夾著尾巴逃跑了，損失體力<font color=red>" & xqz & "</font>點，防御損失<font color=red>" & 2*xqz & "</font>點！"
end if
else
conn.execute "update 用戶 set 操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 體力=體力-"  & xqz & " where 姓名='" & to1 & "'"
rs.close
rs.open "select * from 用戶 where 姓名='" & to1 & "'",conn
if rs("體力")<-100 then
conn.execute "update 用戶 set 殺人數=殺人數+1 where id="&info(9)
conn.execute "update 用戶 set 狀態='死' where 姓名='" & to1 & "'"
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',now(),'" & info(0) & "','寵物攻擊')"
e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦"
call boot(to1)
end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出--[普通攻擊]--攻擊<font color=9FDF70>" & to1 & "</font>，"&aa
end if
end if
case "伙伴攻擊"
randomize timer
xqz=int(rnd*50)+10
rs.close
sql="SELECT * FROM 寵物狀態 WHERE user='" & to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>使出的伙伴攻擊失敗，因為對方沒有寵物。"
else
id=rs("類型id")
tosay=rs("說明")
petname1=rs("名字")
kills=int(rs("防御")+rs("攻擊")-fangyu-att-(ap+zou)/10)
if kills<10 then
kills=xqz
end if
conn.execute "update 寵物狀態 set 經驗=經驗+10 where user='" & from1 & "'"
conn.execute "update 寵物狀態 set 體力=體力-"  & kills & " where user='" & to1 & "'"
e=""
if rs("體力")-kills<=0 then
conn.execute "update 寵物狀態 set 經驗=經驗+10 where user='" & from1 & "'"

conn.execute "update 寵物狀態 set 體力=體力-"  & kills & " where user='" & to1 & "'"

e="，<font color=red>" & to1 & "</font>的寵物<img src='../myhome/sheep/image/"&tosay&".gif' width=49 height=46>[" & petname1 & "]慢慢的倒了下去  從此江湖又少了一隻寵物！"

end if
pet="<a href=javascript:parent.sw('[" & from1 & "]'); target=f2><font color=red>"&from1&"</font></a>召喚出<img src='../myhome/sheep/image/"&say&".gif' width=49 height=46>[" & lx & "]使出--[伙伴攻擊]--攻擊<font color=red>" & to1 & "</font>的寵物，咬傷<img src='../myhome/sheep/image/"&tosay&".gif' width=49 height=46>[" & petname1 & "]體力<font color=red>" & kills &"</font>"& e
if rs("體力")<=0 then
conn.execute "Delete * From 寵物狀態 Where user='" & to1 & "'"
conn.execute "Delete * From 道具列表 Where 擁有者='" & to1 & "'"
end if
end if
End Select
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<font color=9FDF70>【寵物攻擊】"& pet &"</font>"& t
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜，您的寵物已經發招完成！');location.href = 'javascript:history.go(-1)';</script>"

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
