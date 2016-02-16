<%
'發射子彈
function fashezi(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('發射子彈對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以發射子彈！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 體力,門派 FROM 用戶 WHERE 姓名='"&toname&"'",conn
tl=rs("體力")
money=int(rs("體力")/5)
if rs("門派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失敗，你不能對高級管理員或站長特封的人員操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE 物品名='精制手槍'  and 數量>0 and 擁有者="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有精制手槍嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE 物品名='子彈'  and 數量>0 and 擁有者="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有子彈嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 物品 set 數量=數量-1 where 物品名='子彈'  and 擁有者="&info(9)
randomize()
rs.close
rs.open "select id FROM 用戶 WHERE id="&info(9),conn
'rnd1=int(rnd()*3)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update 用戶 set 體力=體力-"&money&" where 姓名='"&toname&"'"
fashezi=info(0) & "從腰間撥出一把精制手槍<bgsound src=wav/Phant012.wav loop=1>，惡狠狠地對著<font color=red>"&towho&"</font>就是一槍,"&towho&"啊的一聲，體力失去1/5，共計"&money&"點......" 
if tl<=-100 then 
conn.execute "update 用戶 set 狀態='死' where 姓名='"&toname&"'"
fashezi=info(0) & "從腰間撥出一把精制手槍<bgsound src=wav/Phant012.wav loop=1>，惡狠狠地對著<font color=red>"&towho&"</font>就是一槍,"&towho&"啊的一聲，由於體力不支，被當場擊斃，好慘啊..." 
 call boot(towho)
end if

case else
fashezi=info(0) & "從腰間撥出一把精制手槍<bgsound src=wav/Phant012.wav loop=1>，惡狠狠地對著<font color=red>"&towho&"</font>就是一槍,怎奈槍法太臭沒打中......"
end select

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>