<%
'水晶球
function shuijingqiu(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('水晶球對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以使用水晶球！');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('官府人員不可以進行此操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("法力")<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select grade FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs("grade")>=6  then
Response.Write "<script language=JavaScript>{alert('失敗，你不能對高級管理員或站長特封的人員操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE 物品名='水晶球' and 數量>5 and 擁有者="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你5個以上的水晶球嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
rs.close
rs.open "select id FROM 用戶 WHERE id="&info(9),conn
conn.execute "update 用戶 set 法力=法力-1000,操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 登錄=now()+1/564,狀態='眠' where 姓名='"&toname&"'"
conn.execute "update 用戶 set lastkick='"&info(0)&"' where 姓名='"&toname&"'"
shuijingqiu=info(0) & "從<bgsound src=wav/Ye150.wav loop=1>口袋裡拿出水晶球，對著"&towho&"口中念念有詞，隻見水晶球奇幻的散發出光芒照射在<font color=red>"&towho&"</font>的身上，"&towho&"迷迷糊糊的睡著了." 
call boot(towho)

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>