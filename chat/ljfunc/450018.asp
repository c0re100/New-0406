<%
'生日蛋糕
function shendangao(fn1,to1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得100點啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-100 where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='生日蛋糕' and 數量>2 and 擁有者=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有2個生日蛋糕嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 物品 set 銀兩=200000,數量=數量-2 where 物品名='生日蛋糕' and 擁有者=" & info(9)
conn.execute "update 用戶 set 體力=體力+1000000 where id="&info(9)
shendangao=info(0) & "哼哼地從地上撿起一把雪飲狂刀把自己的肚子剖了開來塞了一塊法器生日蛋糕<img src='pic/dz59.gif'>進去，嘿嘿，來吧，打我呀，有本事打死我啊，<font color=red>["&info(0)&"]</font>的體力暴漲100萬點." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>