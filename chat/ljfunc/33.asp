<%'轉賬
function zzyin(fn1,to1,toname)
fn1=abs(fn1)
if fn1<10000 or fn1>5000000 then
	Response.Write "<script language=JavaScript>{alert('轉賬最少1萬，最多500萬！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('轉賬對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以轉賬！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 存款 from 用戶 where id="&info(9),conn
money=rs("存款")
if rs("存款")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你有那麼多的存款嗎？！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
moneyok=money+fn1
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('對方的存款太多就快超過了20億了，為防止丟錢現像請您把錢少轉些吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

conn.execute "update 用戶 set 存款=存款-"& fn1 &" where id="&info(9)
conn.execute "update 用戶 set 存款=存款+"& int(fn1*0.9) &" where 姓名='"&toname&"'"
zzyin=info(0) & "把自己江湖的存款:"& fn1 &"兩，轉賬到"& toname &"的銀行名下，"&toname&"向官府交手續費5%此次操作成功！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>