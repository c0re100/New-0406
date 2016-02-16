<%'增送會費
function give(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('增送會費對像錯了,不可對大家/自己使用');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('戰鬥級太低,為了防止作弊,系統要5級或以上才可以增送會貴！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會員費 FROM 用戶 WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1<1  then
	Response.Write "<script language=JavaScript>{alert('系統規定最太要有1個會費才可以增送！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("會員費")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你好像沒有那麼多的會費哦！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 會員費=會員費-" & fn1 & " where id="&info(9)
conn.execute "update 用戶 set 會員費=會員費+" & fn1 & " where 姓名='"&toname&"'"
give=info(0) & "把" & fn1 & "個會費送給" & toname &","& toname &"連聲說謝謝!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
