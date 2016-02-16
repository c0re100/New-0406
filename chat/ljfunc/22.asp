<%'踢人
function tiren(to1,toname,fn1)
if info(2)<9 then
	Response.Write "<script language=JavaScript>{alert('管理需要9級的才可以的！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('踢人對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs("grade")<info(2) then
	tiren="<bgsound src=wav/gf.wav loop=1>一陣狂風把" & toname & "<IMG SRC=FQ/dz04.gif>刮出了聊天室【聊管=" & info(0) & "】  原因：《"&fn1&"》"
	call boot(toname)
else
	Response.Write "<script language=JavaScript>{alert('因為他是高級管理員，你不能踢他！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','踢人','"& fn1 &"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>