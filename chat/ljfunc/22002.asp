<%
function jiansu(fn1)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('管理需要10級的才可以的！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select lastip FROM 用戶 WHERE 姓名='"& fn1 &"'",conn
unlockip=rs("lastip")
conn.Execute "DELETE FROM iplocktemp WHERE ip='" & unlockip & "'"
jiansu="<bgsound src=wav/gf.wav loop=1>" & fn1 & "有心改過，給予解鎖IP!(聊管=" & info(0) & ")  ["&fn1&"]"

conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& fn1 &"','踢人','"& fn1 &"解鎖IP')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>