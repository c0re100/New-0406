<%
function grdb(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('雙人賭博對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以雙人賭博！');}</script>"
	Response.End
end if
if fn1<>1 and fn1<>2 and fn1<>3 then
Response.Write "<script language=JavaScript>{alert('輸入錯誤,應該是1-3的數字！');}</script>"
Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 配偶,道德,魅力,銀兩 FROM 用戶 WHERE id="&info(9),conn
if rs("道德")<300 or rs("魅力")<300 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('道德與魅力不夠300，人家不和你賭！');}</script>"
Response.End
end if
if rs("銀兩")<1000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('沒有100萬塊錢，不能賭博！');}</script>"
Response.End
end if
rs.close
rs.open "select 銀兩 FROM 用戶 WHERE 姓名='"&toname&"'" ,conn
if rs("銀兩")<1000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('他沒有100萬塊錢！');}</script>"
Response.End
end if
rs.close
conn.execute "update 用戶 set 銀兩=銀兩-1000000 where id="&info(9)
set rs=nothing
conn.close
set conn=nothing
Application("ljjh_kissly")=toname
Application("ljjh_bingwen")=fn1
grdb="["&info(0)&"]向<bgsound src=wav/Faqian.wav loop=1>{"&toname&"}提出賭博要求，賭注100萬白花花的銀子！<input type=button value='石頭' onClick=window.open('dubo.asp?id="&info(9)&"&db=1','d')><input type=button value='剪子' onClick=window.open('dubo.asp?id="&info(9)&"&db=2','d')><input type=button value='布' onClick=window.open('dubo.asp?id="&info(9)&"&db=3','d')>"
end function
%>