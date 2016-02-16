<%'求妾
function qqq(to1,toname)
if toname="大家" then
	Response.Write "<script language=JavaScript>{alert('求妾對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 二婚,配偶,性別,銀兩,門派 from 用戶 where id="&info(9),conn
sf=rs("二婚")
po=rs("配偶")
xbk=rs("性別")
if info(0)=toname then
	if sf="" or sf="無" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
if info(1)="男" then
		Response.Write "<script language=JavaScript>{alert('你沒有小老婆，無法休掉！');}</script>"
else
Response.Write "<script language=JavaScript>{alert('哎，小老公你也要休呀可是你沒有小老公，無法休掉！');}</script>"
end if
		Response.End
	end if
if rs("銀兩")<5000000 then
	Response.Write "<script language=JavaScript>{alert('你沒有500萬塊，人家不干呀！！');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
	conn.execute "update 用戶 set 銀兩=銀兩-5000000,二婚='無' where id="&info(9)
        conn.execute "update 用戶 set 銀兩=銀兩+5000000,二婚='無' where 姓名='"& sf &"'"
        conn.execute "delete * from 合體技 where 姓名男='" & info(9) & "' or 姓名女='" & sf & "'"
       rs.close
	set rs=nothing
	conn.close
	set conn=nothing
if info(1)="男" then
	qqq=info(0) & "給原來的小老婆<font color=red>"& sf &"</font>交納了500萬塊的分手費，終於把<font color=red>"& sf &"</font>給踹了~~~"
else
qqq=info(0) & "給原來的小老公<font color=red>"& sf &"</font>交納了1000萬塊的分手費，終於把<font color=red>"& sf &"</font>給甩了~~~"
end if
	exit function
end if
if sf=toname or po=toname then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]已經是你的人了！');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        Response.End

end if
if rs("二婚")<>"" and rs("二婚")<>"無" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
if info(1)="男" then
Response.Write "<script language=JavaScript>{alert('想和["& toname&"]好，請與現在的小老婆脫離關繫！');}</script>"
else
Response.Write "<script language=JavaScript>{alert('想和["& toname&"]好，請與現在的小老公脫離關繫！');}</script>"
end if
	Response.End
end if
if rs("銀兩")<5000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
Response.Write "<script language=JavaScript>{alert('你沒有500萬兩["& toname&"]不想和你好！');}</script>"
	
	Response.End
end if
rs.close
rs.open "select 性別,配偶 FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs("性別")=xbk  then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
Response.Write "<script language=JavaScript>{alert('搞什麼啊？江湖不允許存在同性戀的！');}</script>"
	Response.End
end if

if info(1)="男" then
qqq=info(0) & "向"& toname &"交了500萬塊禮金，提出了要娶對方為小老婆，也不知道人家同意不！"
else
qqq=info(0) & "向"& toname &"交了1000萬塊禮金，提出了要讓對方做自己的小老公，也不知道人家同意不！"
end if
Application("ljjh_qqq_sf")=toname
Application("ljjh_qqq_td")=info(0)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
