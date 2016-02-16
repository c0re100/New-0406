<%'送會費
function give(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('送會費對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以送會費！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會費 FROM 用戶 WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1>1000000  then
	Response.Write "<script language=JavaScript>{alert('江湖規定會費不可以超過100萬的！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("會費")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你有那麼多的會費嗎？看好再說！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1<100 then
	Response.Write "<script language=JavaScript>{alert('你太摳了吧，不能少於100會費的！');}</script>"
	conn.close
	set rs=nothing	
	rs.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 會費=會費-" & fn1 & " where id="&info(9)
conn.execute "update 用戶 set 會費=會費+" & fn1 & " where 姓名='"&toname&"'"
give=info(0) & "把" & fn1 & "兩會費送給了" & toname &",這下可把"& toname &"樂的直蹦,連聲說謝謝!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
