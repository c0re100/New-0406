<%'心跳
function xintiao(fn1)
if info(3)<30 then
	Response.Write "<script language=JavaScript>{alert('此操作需要戰鬥等級30，才可以操作！！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 內力 FROM 用戶 WHERE id="&info(9),conn
if rs("內力")<1000 then
	Response.Write "<script language=JavaScript>{alert('需要內力1000才可以操作！你不夠呀！！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 內力=內力-1000 where id="&info(9)
xintiao="<marquee height=50 behavior=alternate loop=200 direction=up> 心跳心語 " & fn1 & "(" & info(0) & ")" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>