<%'收徒
function stu(to1)
if to1="大家" or to1=info(0) then
	Response.Write "<script language=JavaScript>{alert('收徒對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以收徒！');}</script>"
	Response.End
end if
if trim(Application("ljjh_bais_sf"))<> trim(info(0)) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]也沒有想拜你為師！');}</script>"
	Response.End
end if
if trim(Application("ljjh_bais_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]也沒有想拜你為師！');}</script>"
	Response.End
end if
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 銀兩=銀兩-50000,師傅='"& info(0) &"',師傅交錢='0' where 姓名='"& Application("ljjh_bais_td") &"'"
stu=Application("ljjh_bais_td") & "向"& info(0) &"交納了5萬塊拜師費，又是點頭又是哈腰的，終於求得"& info(0) &"收自己為徒,隻要師傅在，武功大進的！"
Application("ljjh_bais_sf")=""
Application("ljjh_bais_td")=""
conn.close
set conn=nothing
end function
%>