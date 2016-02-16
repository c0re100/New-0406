<%'單挑
function diantiao(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('單挑對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以單挑！');}</script>"
	Response.End
end if
f=Minute(time())
if f<30 then
	Response.Write "<script language=JavaScript>{alert('現在不能單挑！單挑時間為每小時的後30分鐘,例如：17:30開始18:00結束!');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名2 FROM 單挑",conn
if rs("姓名2")<>"無" then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('已經有人在單挑了，請等一會吧或者請求官府人員解除別人單挑！');}</script>"
	Response.End
	end if
rs.close
rs.open "select 門派,體力 FROM 用戶 WHERE id="&info(9),conn
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('官府的不可以跟人單挑啊！');}</script>"
	Response.End
end if
if rs("體力")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的體力太低了，至少也得3000點以上才能跟人單挑！');}</script>"
	Response.End
end if
rs.close
rs.open "select 等級,門派,體力 FROM 用戶 WHERE 姓名='"&toname&"'" ,conn
if rs("等級")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('比武單挑需要你的等級在10級以上！');}</script>"
	Response.End
end if
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('官府的不可以跟人單挑啊！');}</script>"
	Response.End
end if
if rs("體力")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對方的體力太低了，至少也得3000點以上才能跟人單挑！');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application("dantiao")=toname
diantiao="["&info(0)&"]向{"&toname&"}提出單挑："&fn1&"  <input type=button style='FONT-SIZE: 9pt'  value='接受挑戰' onClick=javascript:;disabled=1;window.open('tiaozhan.asp?id="&info(9)&"&yn=1','d')>"
end function
%>
