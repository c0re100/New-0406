<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
name=Request.Form("name2")
id=int(Request.Form("aaa2"))
money=trim(Request.Form("select"))
select case money
case "a"
money="*2"
case "b"
money="*4"
case "c"
money="*6"
case "d"
money="*8"
case "e"
money="*10"
case "f"
money="*50"
case "g"
money="*100"
case "h"
money="/2"
case "i"
money="/4"
case "j"
money="/6"
case "k"
money="/8"
case "l"
money="/10"
case "m"
money="/50"
case "n"
money="/100"
case else
%>
<script language="vbscript">
alert("超出範圍!")
history.back(1)
</script>
<%
end select
sql="select ID,銀兩 from 商店物品 where 擁有者="&id&" and 物品名='"&name&"' and 數量>0"
set rs1=conn1.execute(sql)
if  rs1.EOF or rs1.bof then
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("沒有該物品!")
history.back(1)
</script>
<%Response.End 
end if
yin=rs1("銀兩")
id=rs1("id")
	sql="update 商店物品 set 銀兩=銀兩"&money&" where id="&id
		conn1.execute sql
set rs1=nothing
conn1.close
set conn1=nothing
Response.Redirect "man.asp"
%>

