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
name=Request.Form("name")
id=int(Request.Form("aaa"))
sql="select 類型,說明,內力,體力,攻擊,防御,堅固度,銀兩,數量,等級 from 物品 where 擁有者="&id&" and 物品名='"&name&"' and 裝備=false and 數量>0"
set rs=conn.execute(sql)
if  rs.EOF or rs.bof then
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("沒有該物品!")
history.back(1)
</script>
<%Response.End 
end if
lx=rs("類型")
say=rs("說明")
nl=rs("內力")
tl=rs("體力")
gj=rs("攻擊")
fy=rs("防御")
dj=rs("等級")
jgd=rs("堅固度")
yin=rs("銀兩")
sl=rs("數量")
sql="delete * from  物品 where 擁有者="&id&" and 物品名='"&name&"'"
set rs=conn.execute(sql)
sql="select id from 商店物品 where 擁有者="&id&" and 物品名='"&name&"' and 數量>0"
rs1.Open sql,conn1,1,1
if  rs1.EOF or rs1.bof then
sql="insert into 商店物品(物品名,擁有者,類型,說明,內力,體力,攻擊,防御,堅固度,數量,銀兩,等級) values ('"&name&"',"&info(9)&",'"&lx&"','"&say&"',"&nl&","&tl&","&gj&","&fy&","&jgd&","&sl&","&yin&","&dj&")"
conn1.execute sql
else
id=rs1("id")
	sql="update 商店物品 set 數量=數量+" & sl & " where id="&id
	conn1.execute sql
end if

set rs1=nothing
conn1.close
set conn1=nothing
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "man.asp"
%>

