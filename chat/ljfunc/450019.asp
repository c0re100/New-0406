<%
'絕情刀
function jiuqingdao(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('操作對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<20 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要20級以上才可以使用絕情刀！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-1000,操作時間=now() where id="&info(9)
rs.close
rs.open "select 體力,門派 FROM 用戶 WHERE 姓名='"&toname&"'",conn
tl=rs("體力")
money=int(rs("體力")/3)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='絕情刀' and 數量>=1 and 擁有者=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有絕情刀嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="& id &" and 擁有者=" & info(9)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update 用戶 set 體力=體力-"&money&" where 姓名='"&toname&"'"
jiuqingdao=info(0) & "從腰間撥出一把絕情刀<bgsound src=wav/fx42.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<font color=red>"&towho&"</font>刺了下去,受死吧，"&towho&"啊的一聲，被刺中心脈,曾經滄海難為水啊，體力失去1/3，共計"&money&"點......" 
if tl<=0 then 
conn.execute "update 用戶 set 狀態='死' where 姓名='"&toname&"'"
jiuqingdao=info(0) & "從腰間撥出一把絕情刀<bgsound src=wav/ZR2199.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<font color=red>"&towho&"</font>刺了下去,受死吧,"&towho&"啊的一聲，由於體力不支，被當場刺死，自古多情空餘恨啊..." 
 call boot(towho)
end if
case else
jiuqingdao=info(0) & "從腰間撥出一把絕情刀<bgsound src=wav/fx42.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<font color=red>"&towho&"</font>刺了下去,受死吧,怎料是愛到深處心恨誰，沒刺到啊......"
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>