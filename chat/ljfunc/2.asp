<%'點穴
function dian(to1,fn1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('點穴對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('想作什麼呀，你可不是官府的！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,姓名 from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('沒有這個人？你是不是看錯了！');}</script>"
	
	Response.End
end if
if rs("grade")<info(2)  then
	call boot(toname)
	conn.execute "update 用戶 set 登錄=now()+0.02,狀態='點穴' where 姓名='"&toname&"'"
	dian=info(0)& "對" & toname & "使用了點穴術，" & toname & "呆呆地不動了  需28分鐘才可以上線  "
else
	dian=info(0) & "對" & toname & "使用了點穴術，可是他是官府中高級管理員！"
end if
'記錄
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','點穴','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
