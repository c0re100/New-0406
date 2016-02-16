<%'傳授
function cuan(fn1,to1,toname)
if toname="" or toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('傳授對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以傳內力，傳內力最少需要10級！');}</script>"
	Response.End
end if
fn1=abs(fn1)
if fn1<200 then
Response.Write "<script language=JavaScript>{alert('傳送內力一次最少200的，別太小氣！');}</script>"
Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 內力 FROM 用戶 WHERE id="&info(9),conn
if fn1>30000  then
	Response.Write "<script language=JavaScript>{alert('傳內力不要大於3000好不！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>3000 then
	if info(4)=0 then 
	Response.Write "<script language=JavaScript>{alert('非會員傳內力不能超過3000,會員不能超過1000000');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	else
		if fn1>1000000 then
		
	Response.Write "<script language=JavaScript>{alert('非會員傳內力不能超過3000,會員不能超過1000000');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
		end if
end if
end if
if rs("內力")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪裡那麼多的內力？搞錯了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 內力=內力-" & fn1 & " where id="&info(9)
conn.execute "update 用戶 set 內力=內力+" & fn1 & " where 姓名='"&toname&"'"
cuan=info(0) & "發功，滿臉通紅，呼呼直喘,頭上青煙直冒，終於把" & fn1 & "的內力傳給了" & towho &"了，"&  towho  &"萬分感謝！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>