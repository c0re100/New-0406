<%'主題
function titl(fn1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('心動在線，需要戰鬥等級10級，你不夠呀！！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 內力,等級 FROM 用戶 WHERE id=" & info(9),conn
if rs("內力")<1000 or rs("等級")<10 then
	Response.Write "<script language=JavaScript>{alert('需要內力1000、等級10才可以，好好練幾天吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
else
	conn.execute "update 用戶 set 內力=內力-1000 where id=" & info(9)
	titl="<marquee border=1  scrolldelay=175><img src=dog.gif>" & fn1 & "【"&info(0)&"】</marquee>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>