<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
to1=request.form("to1")
gw=left(to1,1)
if IsNumeric(gw)=false and Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
zz=split(to1,"級")
gw=zz(0)
	Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if Application("jingda")=1 then
	Response.Write "<script Language=Javascript>alert('目前由於人少的原因系統被設置為禁打狀態，要想打架請找官府人員開啟！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
xx1=rs("姓名1")
xx2=rs("姓名2")
rs.close
rs.open "select 門派,狀態,武功,攻擊,銀兩,防御,體力,殺人數,保護,體力,操作時間,等級,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6 FROM 用戶 WHERE id="&info(9),conn
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
lxjwg1=rs("武功")
lxjgj1=rs("攻擊")
lxjfy1=rs("防御")
lxjyl1=rs("銀兩")
egj1=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")
efy1=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('請取消練功保護再操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("狀態")="死" or rs("體力")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if lxjyl1<=100000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你沒有100000000兩銀子，不能發出這招特技，還是多賺點錢吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("門派")="出家" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經出家了，不可以再殺人了！');location.href = 'javascript:history.go(-1)';</script>"
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
	Response.Write "<script Language=Javascript>alert('此絕招需要戰鬥等級5級以上才可以操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<100 then
	s=100-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('警告：請等"& s &"秒再發招,可別累著！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gw=left(to1,1)
bearcc=12345
if bearcc=12345 then
if IsNumeric(gw)=true then
select case info(8)
case "戰士"
lxjgj1=lxjgj1*1
case "銅盔戰士"
lxjgj1=lxjgj1*2
case "銀鎧戰士"
lxjgj1=lxjgj1*3
case "金甲戰士"
lxjgj1=lxjgj1*4
case "神龍戰士"
lxjgj1=lxjgj1*5
case "勇士"
lxjfy1=lxjfy1*1
case "百戰勇士"
lxjfy1=lxjfy1*2
case "虎威勇士"
lxjfy1=lxjfy1*3
case "擒龍勇士"
lxjfy1=lxjfy1*3
case "金剛勇士"
lxjfy1=lxjfy1*4
case "道士"
lxjwg1=lxjwg1*1
case "水道士"
lxjwg1=lxjwg1*1.5
case "水法師"
lxjwg1=lxjwg1*2
case "水真人"
lxjwg1=lxjwg1*2.5
case "水天師"
lxjwg1=lxjwg1*3
end select
zz=split(to1,"級")
gw=zz(0)
if info(3)>gw*2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('[以你目前的等級用不著打比你低2倍等級的怪物了]');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
  end if
rs.close
rs.open "select * from 怪物區 where 怪物='"&to1&"'",conn
gwgj=rs("攻擊")
gwfy=rs("防御")
gwtl=rs("體力")
killer=int((lxjgj1+lxjwg1+egj1+100000000)-gwfy*2)
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
kaa=int(rnd*99)+1
killer=180000000
killer=killer+kaa
end if
killergw=int(gwgj-(lxjwg1+efy1+lxjfy1))
if killer<=1 then
killer=int(rnd*99)+1
end if
if killergw<=1 then
killergw=int(rnd*99)+1
end if
gwclick="<font color=6FBFEF><a href=javascript:parent.sw('["+gw+"級怪物]'); target=f2>"&aa&""& to1 & "</a></font>"
meclick="<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]'); target=f2>["&info(0)&"]</a>"
aa="<img src='../ico/"+gw+".gif'  border='0' width='36' height='36'>"
bb="<img src='../ico/"& info(10) &"-2.gif' width='36' height='36'>"
conn.execute "update 怪物區 set kill=kill-"  & killer & " where 怪物='"&to1&"'"
conn.execute "update 用戶 set 銀兩=銀兩-100000000 where id="&info(9)
rs.close
rs.open "select * from 怪物區 where 怪物='"&to1&"'",conn
if rs("kill")<0 then
jingye=int(rs("等級"))
if rs("殺人數")>10 then
sj=DateDiff("n",rs("操作時間"),now())
if sj<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('此怪物已經死亡，請你等上["& ss &"]分鐘,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
conn.execute "update 怪物區 set 殺人數=0 where 怪物='"&to1&"'"
end if
conn.execute "update 怪物區 set kill=體力,殺人數=殺人數+1,操作時間=now() where 怪物='"&to1&"'"
conn.execute "update 用戶 set allvalue=allvalue+"&jingye&" where 姓名='"&info(0)&"'"
stunt=info(0) & "<bgsound src=wav/jiuzhao3.wav loop=1><img src='img/18.gif'>運足內力對"&aa&""&gwclick&"使用了江湖上明月宮主留傳下來的曠世絕學<font color=6FBFEF>明月神功</font>，隻見光芒萬丈，一束強光發射出來殺傷"&gwclick&"生命值：" & killer & "點,"&gwclick&"終於不支倒地，慘死在"&meclick&"手上！"&meclick&"獲得"&jingye&"點經驗值！"
call jiujian(stunt)
else
stunt=info(0) & "<bgsound src=wav/jiuzhao3.wav loop=1><img src='img/18.gif'>運足內力對"&aa&""&gwclick&"使用了江湖上明月宮主留傳下來的曠世絕學<font color=6FBFEF>明月神功</font>，隻見光芒萬丈，一束強光發射出來殺傷"&gwclick&"生命值：" & killer & "點"&gwclick&"奮起反擊咬傷"&meclick&"生命值:<font color=red>"&killergw&"</font>點!"
conn.execute "update 用戶 set 體力=體力-"  & killergw & ",操作時間=now() where 姓名='"&info(0)&"'"
call jiujian(stunt)
rs.close
rs.open "select 體力 from 用戶 where 姓名='"&info(0)&"'",conn
if rs("體力")<0 then
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
sd(196)="6FBFEF"
sd(197)="6FBFEF"
sd(198)="對"
sd(199)="<font color=6FBFEF>【系統】["&info(0)&"]</font>大俠武學造詣未成，慘死在"&gw&"級怪物手上!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock

Response.Write "<script language=javascript>alert('【"&info(0)&"】你學藝不精，慘死在怪物手上了！');</script>"
if info(4)=0 then 
call boot(info(0))
end if
end if
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜，您的明月神功已經發招完成！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
rs.close
rs.open "select id,門派,會員等級,武功,攻擊,殺人數,保護,防御,狀態,體力,grade,等級,b1,b2,b3,b4,b5,b6 FROM 用戶 WHERE 姓名='" & to1 & "'",conn
menpai=rs("門派")
jhhy=rs("會員等級")
lxjwg2=rs("武功")
lxjgj2=rs("攻擊")
lxjfy2=rs("防御")
idd=rs("id")
r=rs("等級")
efy2=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
killnum=rs("殺人數")
if xx1=to1 or xx2=to1 then
 if diantiaook=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不可以打正在單挑中的人!');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
end if
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('對方正在閉關！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("狀態")="死" or rs("體力")<-100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('人都已經死了還打啊');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("門派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & to1 & "]他已經出家，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("等級")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & to1 & "]等級太低，經不起你打啊！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei=100000000


killer=int((lxjwg1+lxjgj1+egj1+(denji1/r)*500000)-(lxjwg2+lxjfy2+efy2))
if killer<=100 then
randomize timer
killer=int(rnd*99)+1
end if
conn.execute "update 用戶 set 體力=體力-"  & killer & " where 姓名='" & to1 & "'"

conn.execute "update 用戶 set 銀兩=銀兩-100000000 where id="&info(9)

e=""
rs.close
rs.open "select 體力 FROM 用戶 WHERE 姓名='" & to1 & "'",conn
if rs("體力")<0 then
conn.execute "update 用戶 set 殺人數=殺人數+1,kill=kill+1,allvalue=allvalue+"&r&" where id="&info(9)
conn.execute "update 用戶 set 內力=100,武功=100,攻擊=100,防御=100,狀態='死',操作時間=now(),lastkick='"&info(0)&"',allvalue=allvalue-"&r&" where 姓名='" & to1 & "'"
'conn.execute "delete * from 物品 where 擁有者="&ask&" and 類型<>'卡片' and 類型<>'法器' and 裝備=false"
if xx1=info(0) or xx2=info(0) then
conn.execute "update 單挑 set 姓名1='無',姓名2='無'"
end if
e1=","

rs.close
rs.open "select id,物品名 from 物品 where ( 類型='右手' or 類型='左手' or 類型='雙腳' or 類型='頭盔' or 類型='裝飾' or 類型='盔甲' ) and 數量>0 and 堅固度>0 and 裝備=1 and 擁有者="&idd,conn
if not rs.eof or not rs.bof then
id=rs("id")
zb=rs("物品名")
randomize timer
r=int(rnd*r)
conn.execute "update 物品 set 堅固度=堅固度-"&r&" where id="&id
rs.close
rs.open "select 堅固度 from 物品 where id="&id,conn
if rs("堅固度")<0 then 
call jiechu(idd,id)
end if
e1=",["&to1&"]的裝備["&zb&"]損壞"&r&"點!"
end if
if killnum>=10 then
rs.close
rs.open "select id,物品名,說明 from 物品 where ( 類型='右手' or 類型='左手' or 類型='雙腳' or 類型='頭盔' or 類型='裝飾' or 類型='盔甲' ) and 數量>0  and 裝備=1 and 擁有者="&idd,conn
if not rs.eof or not rs.bof then
id=rs("id")
	wupin=rs("物品名")
	sm=rs("說明")

Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',getdate(),'" & info(0) & "','明月神功')"
e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦,[" & info(0) & "]戰鬥經驗增加"&r&"點," & to1 & "損失戰鬥經驗"&r&"點!" & to1 & "一身功力盡廢，死在江湖....掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' width='32' height='32'  align='top' alt='"&wupin&"'></a>" 
call jiechucolor(idd,id)
else
 e="點"&e1&"" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦,[" & info(0) & "]戰鬥經驗增加"&r&"點," & to1 & "損失戰鬥經驗"&r&"點!"
end if
else
 rs.close
rs.open "select id,物品名,說明 from 物品 where  數量>0 and 裝備=false and 類型<>'卡片' and 擁有者="&idd,conn
if rs.eof or rs.bof then
           e="點"&e1&"" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦,[" & info(0) & "]戰鬥經驗增加"&r&"點," & to1 & "損失戰鬥經驗"&r&"點!" & to1 & "一身功力盡廢，死在江湖...."
else
id=rs("id")
	wupin=rs("物品名")
	sm=rs("說明")

Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & to1 & "',getdate(),'" & info(0) & "','明月神功')"
e="點，" & to1 & "慢慢的倒了下去  從此江湖上又少了一隻大蝦,[" & info(0) & "]戰鬥經驗增加"&r&"點," & to1 & "損失戰鬥經驗"&r&"點!"&e1&"" & to1 & "一身功力盡廢，死在江湖....掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' width='32' height='32'  align='top' alt='"&wupin&"'></a>" 
end if
end if
if info(4)=0 then 
call boot(to1)
end if

end if
stunt=info(0) & "運足內力對<img src='img/18.gif'><font color=6FBFEF>" & to1 & "</font><bgsound src=wav/jiuzhao3.wav loop=1>使用了江湖上明月宮主留傳下來的曠世絕學<font color=6FBFEF>明月神功</font>，隻見光芒萬丈，一束強光發射出來殺傷" & killer & e
call jiujian(stunt)
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜，您的明月神功已經發招完成！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
else
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('錯誤：您正在使用的靈劍江湖不是正版，請支持正版江湖，靈劍總站：http://xajh.us');location.href = 'javascript:history.go(-1)';</script>"
end if
sub jiujian(stunt)
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
sd(195)=info(0)
sd(196)="6FBFEF"
sd(197)="6FBFEF"
sd(198)="對"
sd(199)="<font color=6FBFEF>【明月神功】"& stunt &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub 

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
sub jiechu(idd,id)
rs.close
rs.open "SELECT 類型 FROM 物品 WHERE 擁有者=" & idd & " and 數量>0 and id=" & id,conn
leixing=rs("類型")
select case leixing
case "頭盔"
conn.execute "update 用戶 set a1=0,b1=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "左手"

conn.execute "update 用戶 set a2=0,b2=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "右手"

conn.execute "update 用戶 set a3=0,b3=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "雙腳"

conn.execute "update 用戶 set a4=0,b4=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "裝飾"

conn.execute "update 用戶 set a5=0,b5=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "盔甲"

conn.execute "update 用戶 set a6=0,b6=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

end select
conn.execute "delete  from  物品  where id="& id
end sub
sub jiechucolor(idd,id)
rs.close
rs.open "SELECT 類型 FROM 物品 WHERE 擁有者=" & idd & " and 數量>0 and id=" & id,conn
leixing=rs("類型")
select case leixing
case "頭盔"
conn.execute "update 用戶 set a1=0,b1=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "左手"

conn.execute "update 用戶 set a2=0,b2=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "右手"

conn.execute "update 用戶 set a3=0,b3=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "雙腳"

conn.execute "update 用戶 set a4=0,b4=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "裝飾"

conn.execute "update 用戶 set a5=0,b5=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

case "盔甲"

conn.execute "update 用戶 set a6=0,b6=0 where id=" & idd
conn.execute "update 物品 set 裝備=false where id=" & id

end select
end sub
%>