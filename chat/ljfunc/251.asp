<%'下山
function xiashan(to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 師傅,銀兩,等級 from 用戶 where id="&info(9),conn
sf=rs("師傅")
	if sf="" or sf="無" then
		Response.Write "<script language=JavaScript>{alert('你沒有師傅，無法脫離師門！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
if rs("銀兩")<80000 then
	Response.Write "<script language=JavaScript>{alert('你沒有8萬塊，人家不干呀！！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	conn.execute "update 用戶 set 銀兩=銀兩-50000,師傅='無',保留1='保留' where id="&info(9)
	xiashan=info(0) & "向原師傅"& sf &"交納了8萬塊的分手費，終於與"& sf &"脫離了師徒關繫！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
