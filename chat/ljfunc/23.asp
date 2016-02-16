<%
function xindong(fn1)
if info(3)<15 then
	Response.Write "<script language=JavaScript>{alert('心動在線，需要戰鬥等級15級，你不夠呀！！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 內力 FROM 用戶 WHERE id="&info(9),conn
if rs("內力")<1000 then
	Response.Write "<script language=JavaScript>{alert('心動：需內力1000，你不夠呀！！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
else
	conn.execute "update 用戶 set 內力=內力-1000 where id="&info(9)
	xindong="<marquee width=100% behavior=alternate scrollamount=10><img src=p71.gif>" & fn1 & "【" & info(0) & "】" & "</marquee>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>