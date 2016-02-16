<%
'送法力
function songfali(fn1,to1,toname)
fn1=abs(fn1)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('傳送法力對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以傳送法力！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力 FROM 用戶 WHERE id="&info(9),conn
if Application("ljjh_dubozhang2")=info(0) then
Response.Write "<script language=JavaScript>{alert('你現在是法力賭局的莊家不可以傳別人法力！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<fn1 then
Response.Write "<script language=JavaScript>{alert('你有那麼多的法力嗎？！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fn1>50000  then
	if info(4)=0 then 
	Response.Write "<script language=JavaScript>{alert('非會員傳法力不能超過5萬,會員不能超過10萬');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	else
	if fn1>100000 then
	Response.Write "<script language=JavaScript>{alert('非會員傳法力不能超過5萬,會員不能超過10萬');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
end if
conn.execute "update 用戶 set 法力=法力-"& fn1 &" where id="&info(9)
conn.execute "update 用戶 set 法力=法力+"& fn1 &" where 姓名='"&toname&"'"
songfali=info(0) & "把自己的"& fn1 &"點法力傳送給了<font color=red>"& toname&"</font>，"&toname&"法力大增！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function

%>