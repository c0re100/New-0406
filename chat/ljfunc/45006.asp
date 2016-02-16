<%
'魔力鑽石
function molishi(fn1,to1)
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
rs.open "select id FROM 物品 WHERE 物品名='魔力鑽石' and 數量>0 and 擁有者=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有魔力鑽石嗎？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="& id &" and 擁有者=" & info(9)
conn.execute "update 用戶 set 法力=法力+2000 where id="&info(9)
molishi=info(0) & "從口袋裡拿出魔力鑽石，輕輕一擦,魔力鑽石金光閃閃，傾刻間<font color=red>"&info(0)&"</font>的法力暴漲2000點." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>