<%'送會員
function give(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('送會員,不可對大家/自己使用');}</script>"
	Response.End
exit function
end if
if info(2)<=10 then%>
	Response.Write "<script language=JavaScript>{alert('你管理級沒有11,不能送會員給玩家！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會員等級 FROM 用戶 WHERE id="&info(9),conn
fn1=int(abs(fn1))
end if
conn.execute "update 用戶 set 會員等級=會員等級-0 where id="&info(9)
conn.execute "update 用戶 set 會員費=會員費+" & fn1 & " where 姓名='"&toname&"'"
give=info(0) & "把" & fn1 & "級會員送給" & toname &","& toname &"連聲說謝謝!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
