<%'獎勵
function jiangli(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('獎勵對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<5 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要5級以上才可以獎勵！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,grade,身份 FROM 用戶 WHERE id="&info(9),conn
menpai=rs("門派")
if rs("grade")<5 and (rs("身份")<>"堂主" or rs("身份")<>"護法" or rs("身份")<>"長老" or rs("身份")<>"掌門" or rs("身份")<>"副掌門") then
	Response.Write "<script language=JavaScript>{alert('想作什麼呀，你是什麼身份呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")=5 and (rs("身份")="掌門" or rs("身份")="副掌門") then
	fn1=int(abs(fn1))
if fn1>100000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('獎勵不可以超過1億！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=4 and rs("身份")<>"長老" then
	fn1=int(abs(fn1))
if fn1>10000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('獎勵不可以超過1000萬！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=3 and rs("身份")<>"護法" then
	fn1=int(abs(fn1))
if fn1>1000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('獎勵不可以超過100萬！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=2 and rs("身份")<>"堂主" then
	fn1=int(abs(fn1))
if fn1>500000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('獎勵不可以超過50萬！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
rs.close
rs.open "select 幫派基金 FROM 門派 WHERE 門派='" & menpai & "'",conn
if rs("幫派基金")<fn1 then
	Response.Write "<script language=JavaScript>{alert('幫派裡只有："&rs("幫派基金")&"兩,哪裡有那麼多的錢呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select 門派 FROM 用戶 WHERE 姓名='"&toname&"'",conn
if trim(rs("門派"))<> trim(menpai) then
	Response.Write "<script language=JavaScript>{alert('搞錯了吧，["&toname&"]並不是你派弟子！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 門派 set 幫派基金=幫派基金-" & fn1 & " where 門派='" & menpai & "'"
conn.execute "update 用戶 set 銀兩=銀兩+" & fn1 & ",門派基金=門派基金-"& fn1 &" where 姓名='"&toname&"'"
jiangli=info(0) & "把自己門派["&menpai&"]的基金發給弟子" & toname &","& fn1 & "兩作為獎勵！"& toname &"樂的直蹦,連聲說謝謝!"
fn1="門派："&menpai&"  銀兩："&fn1
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','獎勵','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>