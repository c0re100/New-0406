<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=Request.Form("name")
id=Request.Form("shopname")
num=int(abs(Request.form("num")))
shang=Request.Form("shang")
money=int(abs(Request.form("money")))
money=money*num

%><!--#include file="data.asp"--><%
sql="SELECT 店主 FROM 商店"
rs1.open sql,conn1,1,1
user=rs1("店主")
if user=info(0) then 
%>
<script language="vbscript">
alert("別亂來，怎麼可以自己買賣東西呀！")
history.back(1)
</script>
<%Response.End 
end if
sql="select 銀兩 from 用戶 where id="&info(9)
set rs=conn.Execute (sql)
if rs("銀兩")<money then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("您的銀兩不夠，攢點錢再來吧！")
history.back(1)
</script>
<%Response.End 
end if
sql="select 類型,數量,說明,內力,體力,攻擊,防御,堅固度,等級,銀兩,sm from 商店物品 where 擁有者="&id&" and 物品名='"&name&"'"
set rs1=conn1.execute(sql)
if rs1("數量")<int(num) then
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("商店存貸不足！")
history.back(1)
</script>
<%Response.End 
end if
lx=rs1("類型")
nl=rs1("內力")
tl=rs1("體力")
gj=rs1("攻擊")
fy=rs1("防御")
dj=rs1("等級")
jgd=rs1("堅固度")
yin=rs1("銀兩")
say=rs1("說明")
sm=rs1("sm")
sql="update 用戶 set 銀兩=銀兩-"&money&" where id="&info(9)
conn.execute(sql)
sql="update 商店 set 注冊資金=注冊資金+"&money&" where 商店名='"&shang&"'"
conn1.execute(sql)

sql="update 商店物品 set 數量=數量-"&num&" where 擁有者="&id&" and 物品名='"&name&"'"
set rs1=conn1.execute(sql)

sql="select id from 物品 where 物品名='"&name&"' and 擁有者=" & info(9)
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
	sql="insert into 物品(物品名,擁有者,類型,說明,sm,內力,體力,攻擊,防御,堅固度,數量,銀兩,等級,會員) values ('"&name&"',"&info(9)&",'"&lx&"','"&say&"',"&sm&","&nl&","&tl&","&gj&","&fy&","&jgd&","&num&","&yin&","&dj&","&aaao&")"
	conn.execute sql
else
id=rs("id")
	sql="update 物品 set 數量=數量+" & num & " where id="&id
	conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("購買成功！")
location.href = "javascript:history.go(-1)"
</script>
