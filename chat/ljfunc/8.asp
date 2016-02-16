<%'斬首
function zhanshou(fn1,to1)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('斬首需要10級以上管理員操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,grade from 用戶 where id="&to1,conn
toname=rs("姓名")
if info(2)<rs("grade") then
	Response.Write "<script language=JavaScript>{alert('你不可以對高級管理員操作！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 狀態='死',grade=1,門派='遊俠',武功=0,防御=0,攻擊=0,道德=0,魅力=0 where id="&to1
call boot(toname)
zhanshou= info(0) &":<font color=red>官府的人喊到:人犯" & toname & "斬立決!</font>"& fn1
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & toname & "','斬首',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>