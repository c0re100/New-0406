<%'拼命
function pingmin(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('拼命對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<5 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要5級以上才可以拼命！');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('官府人員不可操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 體力 FROM 用戶 WHERE id="&info(9),conn
if rs("體力")<500 then
	Response.Write "<script language=JavaScript>{alert('你的體力小於500呀，沒力氣和人家拼！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select grade,體力 from 用戶 where 姓名='"&toname&"'",conn
if rs("grade")>=6 then
	Response.Write "<script language=JavaScript>{alert('你想作啥找官府拼命啊？！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
tl=rs("體力")/2
randomize()
r=int(rnd*6)+1
select case r
case 2
conn.execute "update 用戶 set 體力="&tl&" where 姓名='"&toname&"'"
pingmin=info(0)& "在身上綁了捆炸藥衝向<bgsound src=wav/7.wav loop=1>["&towho&"]喊道:我跟你拼啦,呀呀呀,隻聽一聲爆炸聲，["&info(0)&"]全身炸得粉碎,["&towho&"]躲的快,幸免於難體力損失一半！"
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & info(0) & "',now(),'" & info(0) & "','同歸於盡')"
case 3
pingmin=info(0)& "在身上綁了捆炸藥大吼大叫的跑到<bgsound src=wav/7.wav loop=1>[" & towho & "]跟前說喊道:我跟你拼啦，呀呀呀，隻聽一聲爆炸聲，["&info(0)&"]和["&towho&"]同歸於盡一塊被炸得粉碎！"
conn.execute "update 用戶 set 狀態='死' where 姓名='"&toname&"'"

conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & towho & "',now(),'" & info(0) & "','同歸於盡')"
case else
pingmin=info(0)& "在身上綁了捆炸藥大吼大叫的跑到<bgsound src=wav/7.wav loop=1>["&towho&"]跟前說喊道:我跟你拼啦，呀呀呀，隻聽一聲爆炸聲，["&info(0)&"]全身炸得粉碎,["&towho&"]幸虧閃得快才躲過一劫！"
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & info(0) & "',now(),'" & info(0) & "','同歸於盡')"
end select
conn.execute "update 用戶 set 狀態='死',體力=0,內力=0,武功=0,攻擊=0,防御=0,法力=0 where id="&info(9)
call boot(info(0))
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>