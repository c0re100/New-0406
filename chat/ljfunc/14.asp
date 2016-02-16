<%
function attack(fn1,to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
xx1=rs("姓名1")
xx2=rs("姓名2")
rs.close
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select 武功,內力,級別,類型 FROM 武功 WHERE   門派='" & info(5) & "' and 武功='"&fn1&"' ",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('你的門派:"& menpai &" 並沒有這樣的武功["& fn1 &"]呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
xiuji=rs("級別")
nlin=rs("內力")

if rs("類型")="防御" then
rs.close
rs.open "select 操作時間,狀態,內力 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=30-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒鐘,再操作！');}</script>"
	Response.End
end if
if rs("內力")<nlin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的內力不足["&nlin&"]無法施展！');</script>"
	Response.End
end if
if rs("狀態")="死" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已經死了呀假死？真死？！');</script>"
	Response.End
end if
randomize timer
r=int(rnd*120)+200
xiujikill=xiuji*r
nl=xiujikill*4
conn.execute "update 用戶 set 防御=防御+"&xiujikill&",內力=內力-"&xiujikill&"*100,操作時間=now() where id="&info(9)
attack=xinxi & info(0) & "施展["&info(5)&"]的<font color=red>[防御術]</font><font color=#FFFDAF>"& fn1 &"</font>[<font color=#FFFDAF>"&xiuji&"</font>]級,自身防御功能提升<font color=red>"&xiujikill&"</font>點,內力消耗<font color=red>"&nl&"</font>點!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("類型")="恢復" then
rs.close
rs.open "select 操作時間,內力 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=30-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒鐘,再操作！');}</script>"
	Response.End
end if
if rs("內力")<nlin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的內力不足["&nlin&"]無法施展！');</script>"
	Response.End
end if
randomize timer
r=int(rnd*120)+200
xiujikill=xiuji*r
nl=xiujikill*4
conn.execute "update 用戶 set 體力=體力+"&xiujikill&",操作時間=now() where 姓名='"&toname&"'"
conn.execute "update 用戶 set 內力=內力-"&xiujikill&"*50,操作時間=now() where id="&info(9)
attack=xinxi & info(0) & "施展["&info(5)&"]的<font color=red>[恢復術]</font>"& fn1 &"["&xiuji&"]級,幫助<font color=FFFDAF>["&toname&"]</font>提升體力<font color=red>"&xiujikill&"</font>點,內力消耗<font color=red>"&nl&"</font>點!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("類型")="攻擊" then
if toname="" or toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('發招對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
xiujikill=xiuji*2
rs.close
rs.open "select 攻擊,保護,狀態,武功,防御,性別,內力,體力,門派,殺人數,暴豆時間,grade,等級,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6 from 用戶 where id="&info(9),conn
grademe=int(rs("grade")*rs("等級"))
egj1=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")
efy1=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
if rs("門派")="出家" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經出家了，不可以再殺人了！');</script>"
	Response.End
end if
if rs("狀態")="死" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經死了快去復活吧！');</script>"
	Response.End
end if
if  rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & info(0) & "]你正在保護狀態,要打架先開自己的保護吧！');</script>"
	Response.End
end if
if rs("體力")<300 or rs("內力")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('【"&info(0)&"】你現在的內力或體力小於300請不要進行發招命令操作！');</script>"
	Response.End
end if
if rs("殺人數")>=30 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你作惡太多，超過江湖殺人上限30不能再發招了！');}</script>"	
	Response.End
end if
if rs("內力")<nlin then
	Response.Write "<script language=JavaScript>{alert('對不起，該招式需要"&nlin&"點內力!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
yxjgjfrom=rs("攻擊")
yxjfyfrom=rs("防御")
yxjwgfrom=rs("武功")

xb=rs("性別")
sj=DateDiff("n",rs("暴豆時間"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=FFFDAF>威力爆滿</font>"
	baodoudj=2
end if
wgnl=nlin*xiujikill
nlkill=int(nlin/4)


gw=left(toname,1)
if IsNumeric(gw)=true then
select case info(8)
case "戰士"
yxjgjfrom=yxjgjfrom*1
case "銅盔戰士"
yxjgjfrom=yxjgjfrom*2
case "銀鎧戰士"
yxjgjfrom=yxjgjfrom*3
case "金甲戰士"
yxjgjfrom=yxjgjfrom*4
case "神龍戰士"
yxjgjfrom=yxjgjfrom*5
case "勇士"
yxjfyfrom=yxjfyfrom*1
case "百戰勇士"
yxjfyfrom=yxjfyfrom*2
case "虎威勇士"
yxjfyfrom=yxjfyfrom*3
case "擒龍勇士"
yxjfyfrom=yxjfyfrom*3
case "金剛勇士"
yxjfyfrom=yxjfyfrom*4
case "道士"
yxjwgfrom=yxjwgfrom*1
case "水道士"
yxjwgfrom=yxjwgfrom*1.5
case "水法師"
yxjwgfrom=yxjwgfrom*2
case "水真人"
yxjwgfrom=yxjwgfrom*2.5
case "水天師"
yxjwgfrom=yxjwgfrom*3
end select
zz=split(toname,"級")
gw=zz(0)
if info(3)>gw*2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('[以你目前的等級用不著打比你低2倍等級的怪物了]');</script>"
	Response.End
  end if
if gw>20 and session("nowinroom")=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('好好玩，不要搗亂！');</script>"
	Response.End
end if
rs.close
rs.open "select * from 怪物區 where 怪物='"&toname&"'",conn
gwgj=rs("攻擊")
gwfy=rs("防御")
gwtl=rs("體力")
killer=int((yxjgjfrom+wgnl+egj1+yxjwgfrom-gwfy*2)*(baodoudj+xiujikill))
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
kaa=int(rnd*99)+1
killer=180000000
killer=killer+kaa
end if
killergw=int(gwgj*2-(yxjwgfrom+yxjfyfrom+efy1))
if killer<=1 then
killer=int(rnd*99)+1
end if
if killergw<=1 then
killergw=int(rnd*99)+1
end if

aa="<img src='../ico/"+gw+".gif'  border='0' width='36' height='36'>"
bb="<img src='../ico/"& info(10) &"-2.gif' width='36' height='36'>"
conn.execute "update 怪物區 set kill=kill-"  & killer & " where 怪物='"&toname&"'"
rs.close
rs.open "select * from 怪物區 where 怪物='"&toname&"'",conn
if rs("kill")<-100 then
jingye=int((rs("殺人數")+1)+gw)
if rs("殺人數")>100 then
sj=DateDiff("n",rs("操作時間"),now())
if sj<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('此怪物已經死亡，請你等上["& ss &"]分鐘,再操作！');}</script>"
	Response.End
end if
conn.execute "update 怪物區 set 殺人數=0 where 怪物='"&toname&"'"
end if
conn.execute "update 怪物區 set kill=體力,殺人數=殺人數+1,操作時間=now() where 怪物='"&toname&"'"
conn.execute "update 用戶 set allvalue=allvalue+"&jingye&",內力=內力-"&nlin&" where 姓名='"&info(0)&"'"
randomize()
rnd1=int(rnd*info(3))
//rnd1=5
select case rnd1
case 5
attack=xinxi & bb&info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]"& fn1 &"["&xiuji&"]級攻擊"&aa&"" & towho & ",殺傷" & towho & "生命值:<font color=red>"&killer& "</font>點，"& towho &"慘死在地上，["&info(0)&"]獲得戰鬥經驗增加<font color=FFFDAF>"&jingye&"</font>點!<img src='../hcjs/jhjs/001/2003.gif' border='0' align='top' alt='"&wupin&"'>"
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,會員) values ('流星淚',"&info(9)&",'物品',0,1000,0,0,0,1,10000000,2003,0,"&aaao&")"
else 
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
end if
case else
attack=xinxi & bb&info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]"& fn1 &"["&xiuji&"]級攻擊"&aa&"" & towho & ",殺傷" & towho & "生命值:<font color=red>"&killer& "</font>點，"& towho &"慘死在地上，["&info(0)&"]獲得戰鬥經驗增加<font color=FFFDAF>"&jingye&"</font>點!"
end select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
attack=xinxi & info(0) & "施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/26.gif'>攻擊"&aa&""& toname & ",使得"& toname &"體力下降:<font color=red>"& killer &"</font>點,"& toname &"反攻咬傷["&info(0)&"]生命值:<font color=red>"&killergw& "</font>點!"
conn.execute "update 用戶 set 體力=體力-"  & killergw & ",操作時間=now() where 姓名='"&info(0)&"'"
rs.close
rs.open "select 體力 from 用戶 where 姓名='"&info(0)&"'",conn
if rs("體力")<0 then
conn.execute "update 用戶 set 狀態='死',操作時間=now() where 姓名='"&info(0)&"'"
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=red>【系統】</font><font color=FFFDAF>["&info(0)&"]</font>大俠武學造詣未成，慘死在"&gw&"級怪物手上!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=javascript>alert('【"&info(0)&"】你學藝不精，慘死在怪物手上了！');</script>"
	
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if

rs.close
'對方的判斷
rs.open "select id,姓名,狀態,保護,性別,體力,等級,防御,武功,攻擊,會員等級,門派,grade,b1,b2,b3,b4,b5,b6 from 用戶 where 姓名='"&towho&"'",conn
toname=rs("姓名")
yxjfyto=rs("防御")
yxjwgto=rs("武功")
yxjgjto=rs("攻擊")
jhhy=rs("會員等級")
xb1=rs("性別")
menpai1=rs("門派")
idd=rs("id")
efy2=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
gradeto=int(rs("grade")*rs("等級"))
if rs("等級")<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]是新手要注意保護哇！');</script>"
	Response.End
end if
if  rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]正在保護狀態，另想辦法吧！');</script>"
	Response.End
end if
if  rs("體力")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]已經死了你還打呀！');</script>"
	Response.End
end if
if  rs("grade")>6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]是官府管理人員，如對其有意見請與站長聯繫，不要人身攻擊！');</script>"
	Response.End
end if
if rs("門派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]已經出家，不能攻擊！');</script>"
	Response.End
end if
if xx1=towho or xx2=towho then
 if diantiaook=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不可以打正在單挑中的人!');</script>"
	Response.End
end if
end if
killer=int((yxjgjfrom+wgnl+yxjwgfrom+egj1+grademe-(yxjfyto+yxjwgto+efy2+gardeto))*(baodoudj+xiujikill))

if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update 用戶 set 道德=道德-20,內力=內力-" & abs(nlkill) & ",體力=體力-"& int(killer/10) & " where id="&info(9)
conn.execute "update 用戶 set 體力=體力-"  & killer & " where 姓名='"&towho&"'"
rs.close
rs.open "select 體力 from 用戶 where 姓名='"&towho&"'",conn
if rs("體力")>-100 then
if xb="女" then
	attack=xinxi & info(0) & "施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/26.gif'>攻擊" & toname & ",使得"& toname &"體力下降:<font color=red>-"& killer &"</font>點,自己被"& toname &"的武功震得內力下降:<font color=red>-"& nlkill & "</font>點,體力:<font color=red>-"& int(killer/4) &"</font>點."&info(0)&"戰鬥經驗增加2點!"
else
attack=xinxi & info(0) & "施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/26.gif'>攻擊" & toname & ",使得"& toname &"體力下降:<font color=red>-"& killer &"</font>點,自己被"& toname &"的武功震得內力下降:<font color=red>-"& nlkill & "</font>點,體力:<font color=red>-"& int(killer/4) &"</font>點."&info(0)&"戰鬥經驗增加2點!"
end if
conn.execute "update 用戶 set allvalue=allvalue+2 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
	conn.execute "update 用戶 set 殺人數=殺人數+1,allvalue=allvalue+20 where id="&info(9)
	conn.execute "update 用戶 set 狀態='死',操作時間=now(),lastkick='"&info(0)&"' where 姓名='" & towho & "'"
if xx1=info(0) or xx2=info(0) then
conn.execute "update 單挑 set 姓名1='無',姓名2='無'"
end if
'call boot(toname)
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & toname & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
rs.open "select id,物品名,說明 from 物品 where (類型='藥品') and 數量>0 and 裝備=false and 擁有者="&idd,conn
if rs.eof or rs.bof then
if menpai1<>"通緝犯" then 
if xb1="女" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,"&info(0)&"戰鬥經驗增加20點!"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,"&info(0)&"戰鬥經驗增加20點!"
end if
else
if xb1="女" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,因殺死通緝犯，官府獎勵銀兩500萬!"&info(0)&"戰鬥經驗增加20點"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,因殺死通緝犯，官府獎勵銀兩500萬!"&info(0)&"戰鬥經驗增加20點"
end if
conn.execute "update 用戶 set 銀兩=銀兩+5000000 where id="&info(9)
end if
else
id=rs("id")
	wupin=rs("物品名")
	sm=rs("說明")
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
if menpai1<>"通緝犯" then 
if xb1="女" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,"&info(0)&"戰鬥經驗增加20點!["&toname&"]掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' align='top' border='0' alt='"&wupin&"'></a>"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,"&info(0)&"戰鬥經驗增加20點!["&toname&"]掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' align='top' border='0' alt='"&wupin&"'></a>"
end if
else
if xb1="女" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,因殺死通緝犯，官府獎勵銀兩500萬!"&info(0)&"戰鬥經驗增加20點,["&toname&"]掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' align='top' alt='"&wupin&"'></a>"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>施展["&info(5)&"]的<font color=red>[攻擊術]</font>"& fn1 &"["&xiuji&"]級<img src='img/28.gif'>攻擊" & toname & "，殺傷對方體力"& killer &"點，" & toname & "慢慢倒下..江湖從此又少了一位大蝦,因殺死通緝犯，官府獎勵銀兩500萬!"&info(0)&"戰鬥經驗增加20點,["&toname&"]掉落一些物品，大家快搶!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' align='top'  alt='"&wupin&"'></a>"
end if
conn.execute "update 用戶 set 銀兩=銀兩+5000000 where id="&info(9)
end if
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>