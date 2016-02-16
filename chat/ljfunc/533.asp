<%
function saohuang(to1,toname) 
if toname="大家" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('掃黃對像有錯，請看仔細了！！');}</script>"
Response.End
exit function
end if
if info(1)="男" then
Response.Write "<script language=JavaScript>{alert('掃黃行動暫由女俠執行，男的閃一邊！');}</script>"
Response.End
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("ljjh_usermdb") 
rs.open "select 姓名 FROM 名妓 WHERE 介紹='"&toname&"'",conn 
if rs.eof and rs.bof then 
Response.Write "<script language=javascript>{alert('此人沒有從事過拐賣婦女！');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
conn.execute "update 用戶 set 銀兩=銀兩-5000000 where 姓名='"&toname&"'"
conn.execute "update 用戶 set 銀兩=銀兩+1000000 where 姓名='"&info(0)&"'"
conn.execute "delete * from 名妓 where 介紹='"&toname&"'"
saohuang=info(0) & "和一幫姐妹們帶著大刀闖入怡紅院把["&toname&"]圍起來狠狠的揍了一頓，救出了被["&toname&"]拐賣進去的少女，["&toname&"]被搶掉500萬銀兩，["&info(0)&"]得到100萬." 
rs.close
set rs=nothing
conn.close
set conn=nothing 
end function 
%>