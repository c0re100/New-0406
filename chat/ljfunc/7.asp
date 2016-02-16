<%'解除監禁
function shifang(fn1)
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('解除監禁需要8級以上管理員操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 狀態 from 用戶 where 姓名='"& fn1 &"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('你所輸入的用戶名不存在！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("狀態")<>"監禁" then
	Response.Write "<script language=JavaScript>{alert('["& fn1 &"]並沒有被監禁！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 狀態='正常',登錄=now() where 姓名='" & fn1 & "'"
shifang= info(0) &":<font color=red>官府的人看"& fn1 &"有心改過，表現良好,經大家一至同意給他一次改過自新的機會，將其釋放！</font>"
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & fn1 & "','解除監禁',now(),'有心改過，原諒他這一次！')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>