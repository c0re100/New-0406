<%'幸運風采
function xingyu(fn1)
if Application("ljjh_xinyu")=0 or isempty(Application("ljjh_xinyu")) then
	randomize()
	rnd1=cint(rnd()*999)+1
	Application.Lock
	Application("ljjh_xinyu")=rnd1
	Application.UnLock
end if
fn1=int(abs(fn1))
if fn1<0 or fn1>1000 then
	Response.Write "<script language=JavaScript>{alert('輸入錯誤,幸運風采為：1-1000之間的數！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if fn1=Application("ljjh_xinyu") then
	Application.Lock
	Application("ljjh_xinyu")=0
	Application.UnLock
	conn.execute "update 用戶 set 銀兩=銀兩+5000000 where id="&info(9)
	
	xingyu="★靈劍彩票★：恭喜"&info(0)&"您中了靈劍江湖福利采票，號碼："&fn1&"<img src='img/251.gif'>獎金500萬<img src='img/251.gif'>，大家表示恭喜！"
else

	rs.open "select 銀兩 FROM 用戶 WHERE id="&info(9),conn
	if rs("銀兩")<1000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('沒有錢了，沒辦法了！別買了！');}</script>"
		Response.End
	end if
	
	conn.execute "update 用戶 set 銀兩=銀兩-2000 where id="&info(9)
	
	xingyu="★靈劍彩票★："&info(0)&"購買了采票，為靈劍福利事業作貢獻！號碼："&fn1&"沒有中獎,還有下一次，還有機會！獎金：500萬"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>
