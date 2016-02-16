<%'門規處置
function mengui(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('處置對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,身份 FROM 用戶 WHERE id="&info(9),conn
fn1=int(abs(fn1))
mpai1=rs("門派")
if rs("身份")<>"掌門" and rs("身份")<>"長老" and rs("身份")<>"副掌門" then
	Response.Write "<script language=JavaScript>{alert('你是什麼身份？要長老級以上高手才有此功能！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>50000 then
	Response.Write "<script language=JavaScript>{alert('罰款最多只能5萬？！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select 門派,銀兩,存款 FROM 用戶 WHERE 姓名='"&toname&"'",conn
mpai2=rs("門派")
if mpai1<>mpai2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
  Response.Write "<script language=JavaScript>{alert('對方不是你門派的弟子啊？！');}</script>"
Response.End
end if
if rs("銀兩")>=fn1 or rs("存款")>=fn1 then
if rs("銀兩")>=fn1 then
conn.execute "update 門派 set 幫派基金=幫派基金+"& fn1 &" where 門派='" &  info(5) &"'"
conn.execute "update 用戶 set 銀兩=銀兩-" & fn1 & " where 姓名='"&toname&"'"
mengui="掌門："&info(0) & "因弟子" & toname & "違反門規，按門規處置罰款" & fn1 &"兩,所罰款項放入本門基金......"
end if
if rs("存款")>=fn1 then
conn.execute "update 門派 set 幫派基金=幫派基金+"& fn1 &" where 門派='" &  info(5) &"'"
conn.execute "update 用戶 set 存款=存款-" & fn1 & " where 姓名='"&toname&"'"
mengui="掌門："&info(0)& "因弟子" & toname & "違反門規，按門規處置罰款" & fn1 &"兩,所罰款項放入本門基金......"
end if
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
  Response.Write "<script language=JavaScript>{alert('你的弟子身上沒幾兩銀子啊？！');}</script>"
Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
