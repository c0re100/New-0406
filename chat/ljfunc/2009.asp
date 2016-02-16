<%'黑暗減金術
function ddd(fn1,to1,toname)
if info(2)<20000000 then
	Response.Write "<script language=JavaScript>{alert('需要20000000級以上管理員操作！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,體力,內力 FROM 用戶 WHERE 姓名='"&toname&"'",conn
if info(2)<rs("grade") then
	Response.Write "<script language=JavaScript>{alert('不可以對高級管理員操作!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 金幣=金幣-100000 where 姓名='"&toname&"'"
if info(5)<>"官府" then
ddd= info(0) &"[便衣聊管]:<font color=red><bgsound src=wav/daipu.wav loop=1>由於" & toname & "在江湖中違反規則,官府決定對其動用靈劍刑罰，" & toname & "因此減金幣-100000了!</font>"
else
ddd= info(0) &"[官府人員]:<font color=red><bgsound src=wav/daipu.wav loop=1>由於" & toname & "在江湖中違反規則,官府決定對其動用靈劍刑罰，" & toname & "因此減金幣-100000了!</font>"
end if 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>
