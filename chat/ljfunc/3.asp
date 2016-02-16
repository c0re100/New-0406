<%'解穴
function jie(fn1)
fn1=trim(fn1)
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('你不是官府中人，不能解穴！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 狀態,登錄 from 用戶 where 姓名='" & fn1 & "'",conn
if trim(rs("狀態"))="點穴" and (not rs.eof) and rs("登錄")>now() then
	conn.execute "update 用戶 set 登錄=now(),狀態='正常' where 姓名='" & fn1 & "'"
	jie=info(0)& "對" & fn1 & "使用了解穴術，" & fn1 & "猛然醒過來了  "
else
	jie=info(0)& "對" & fn1 & "使用了解穴術，可是" & fn1 & "並沒有被點穴  "
end if
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& fn1 &"','解穴','123')"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>
