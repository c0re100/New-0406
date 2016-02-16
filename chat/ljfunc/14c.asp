<%
function attackc(fn1,to1)
if towho="大家" or towho=info(0) then
	Response.Write "<script language=JavaScript>{alert('拚命對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
ddkm=info(4)
if ddkm=0 then ddkm=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
xx1=rs("姓名1")
xx2=rs("姓名2")
rs.close
rs.open "select 攻擊,武功,性別,內力,門派,殺人數,保護,暴豆時間,等級 from 用戶 where 姓名='"&info(0)&"'",conn
gcd=rs("等級")
if gcd=0 then gcd=1
if rs("門派")="出家" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經出家了，不可以再殺人了！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("保護")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你閉關修練中..不能打架！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
yxjgjfrom=rs("攻擊")
yxjwgfrom=rs("武功")
nlfrom=rs("內力")
menpai=rs("門派")
xb=rs("性別")
if rs("殺人數")>=30 then
	conn.execute "update 用戶 set 保護=false where 姓名='"&info(0)&"'"
	Response.Write "<script language=JavaScript>{alert('你作惡太多，超過江湖殺人上限"& Application("sjjh_killman") &"不能再發招了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("暴豆時間"),now())
xinxi=""
baodoudj=1
if sj<=60 then
xinxi="<font color=red>威力爆滿</font>"
baodoudj=3
end if
rs.close
'開始武功
rs.open "select 魔法,基本攻擊,消耗法力,等級,施法說明,圖檔 from 使用技能 where 技能='"&fn1&"' and 姓名='" &info(0)&"'",conn
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你的技能裡並沒有這樣的武功["& fn1 &"]呀！');}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
eepic=rs("圖檔")
mofa=rs("魔法")
dengji=rs("等級")
sm=rs("施法說明")
wgnl=rs("消耗法力")*dengji
rs.close
'對方的判斷
rs.open "select 姓名,狀態,性別,保護,門派,銀兩,存款,防御,法力,法力點 from 用戶 where 姓名='"&towho&"'",conn
toname=rs("姓名")
xb1=rs("性別")
aa4=rs("防御")
aa4=int(aa4/10)
aa5=rs("法力")
aa6=rs("法力點")
aa7=int(aa5+aa6)
aa7=int(aa7/10)
menpai1=rs("門派")
aa1=rs("銀兩")
aa2=rs("存款")
aa3=int(aa1+aa2)
aa3=int(aa3/100)
if rs("門派")="出家" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('[" & to1 & "]已經出家，不能攻擊！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("保護")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('["&towho&"]保護中..不能攻擊！');location.href = 'javascript:history.go(-1)';</script>"
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
if rs("狀態")="死" then
Response.Write "<script language=JavaScript>{alert('不行呀;"& towho &"已經掛了你們誰給他復活呀');}</script>"
Response.End
end if
rs.close
randomize timer
kaa=int(rnd*99)+1
killer=int(wgnl*gcd)
killer=killer+kaa
killer=killer-aa4
killer=killer*ddkm
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
killer=180000000
killer=killer+kaa
end if
if killer<=100 then
randomize timer
killer=int(rnd*9999)+1
end if
'對方的判斷
randomize timer
conn.execute "update 用戶 set 體力=體力-"  & killer & " where 姓名='"&towho&"'"
rs.open "select ID,體力,會員等級 from 用戶 where 姓名='"&towho&"'",conn
hyid=rs("ID")
hygd=rs("會員等級")
if rs("體力")>-100 then
randomize
aaa=int(rnd*4)
select case aaa
case 1
attackc=xinxi & info(0) & "<bgsound src=3.WAV loop=2>用"& fn1 &"<img src="& eepic &">攻擊" & towho & "(防御力"&aa4&"),使得"& towho &","&sm&"體力下降:<font color=red>-"& killer &"</font>.殺"& towho &"可獲得銀兩<font color=red><b>"&aa3&"</b>兩</font>,法力<font color=red><b>"&aa7&"</b>點</font>"
case 2
attackc=xinxi & info(0) & "<bgsound src=3.WAV loop=2>用"& fn1 &"<img src="& eepic &">攻擊" & towho & ",(防御力"&aa4&")使得"& towho &","&sm&"體力下降:<font color=red>-"& killer &"</font>.殺"& towho &"可獲得銀兩<font color=red><b>"&aa3&"</b>兩</font>,法力<font color=red><b>"&aa7&"</b>點</font>"
case else
attackc=towho&"對<img src='pic/dz44.gif'>"&info(0)&"大叫道:哈哈,只是給我搔癢癢,再來呀.殺"& towho &"(防御力"&aa4&")可獲得銀兩<font color=red><b>"&aa3&"</b>兩</font>,法力<font color=red><b>"&aa7&"</b>點</font>"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
elseif rs("體力")<-100 and hygd=0 then
conn.execute "update 用戶 set 狀態='死',lastkick='"&info(0)&"',銀兩=0,存款=0,法力=0,法力點=0 where 姓名='"&towho&"'"
conn.execute "update 用戶 set 殺人數=殺人數+2,銀兩=銀兩+"&aa3&",法力=法力+"&aa7&",體力=體力+10000,內力=內力+10000 where id="&info(9)
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & towho & "',now(),'" & info(0) & "','"&fn1&"')"
conn.execute "delete * from 物品 where 擁有者="&hyid
conn.execute "delete * from 交易市場 where 擁有者="&hyid
conn.execute "delete * from 技能列表 where 擁有者='"&towho&"'"
conn.execute "delete * from 使用技能 where 姓名='"&towho&"'"
conn.execute "delete * from 賣面 where  擁有者='"&towho&"'"
rs.close
attackc=xinxi & info(0) & "<bgsound src=wav/11a.wav loop=3><bgsound src=wav/10.wav loop=3>用"& fn1 &"攻擊" & towho & ","&sm&",殺傷力<font color=red><b>-"& killer &"</b></font>只聽一聲慘叫，"& towho &"(防御力"&aa4&")躺在地上說道：啊<img src="& eepic &">我死得好慘啊，然後眼睛一閉與世長辭了.."&info(0)&"因此獲得"& towho &"遺留下的<font color=red><b>"&aa3&"</b>兩</font>銀子,法力<font color=red><b>"&aa7&"</b>點</font>"
call boot(toname)
set rs=nothing
conn.close
set conn=nothing
else
conn.execute "update 用戶 set 狀態='死',lastkick='"&info(0)&"' where 姓名='"&towho&"'"
conn.execute "update 用戶 set 殺人數=殺人數+2 where id="&info(9)
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & towho & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
attackc="<bgsound src=wav/11a.wav loop=3><bgsound src=wav/10.wav loop=3>用"& fn1 &"攻擊" & towho & ","&sm&",殺傷力<font color=red><b>-"& killer &"</b></font>只聽一聲慘叫，"& towho &"(防御力"&aa4&")躺在地上說道：啊<img src="& eepic &">我死得好慘啊，然後眼睛一閉與世長辭了.."&info(0)&"因為打死會員,所以什麼東西都沒得到"
call boot(toname)
set rs=nothing
conn.close
set conn=nothing
end if
end function
%>