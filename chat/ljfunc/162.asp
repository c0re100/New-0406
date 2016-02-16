<%
'送款
function songkuan(fn1,to1,toname)
fn1=abs(fn1)
if fn1<10000 or fn1>50000000 then
	Response.Write "<script language=JavaScript>{alert('轉賬最少1萬，最多5000萬！');}</script>"
	Response.End
end if
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('你無權操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 存款 from 用戶 where id="&info(9),conn
conn.execute "update 用戶 set 存款=存款+" & fn1 & " where 姓名='"&toname&"'"
songkuan=info(0) & "把靈劍江湖銀票<img src='pic/Dz01.gif'>:"& fn1 &"兩送給"& toname &"以做鼓勵!"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>