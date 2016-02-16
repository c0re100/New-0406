<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then  response.redirect "../error.asp?id=425"

shopname=Request.Form("shopname")
name=Request.Form("name")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")
shopmoney=Request.Form("shopmoney")
id=Request.Form("id")
if shopname="" or name="" or shopmoney="" or memo="" or shoptype="請選擇" then 
%><script language="vbscript">
alert("每項必填請返回重填!")
history.back(1)
</script>
<%Response.End 
end if
%><!--#include file="data.asp"--><%
sql="select 存款,id from 用戶 where 姓名='"&name&"'"
set rs=conn.Execute (sql)
if  rs.EOF or  rs.BOF then
set rs=nothing
conn.close
set conn=nothing%>
<script language="vbscript">
alert("江湖裡沒有這個人呀!")
history.back(1)
</script>
<%Response.End 
end if
id=rs("id")
if rs("存款")<int(shopmoney) then
%><script language="vbscript">
alert("注冊資金太多，存款太少!")
history.back(1)
</script>
<%Response.End 
end if
sql="SELECT * FROM 商店 where 商店名='"&shopname&"'"
set rs1=conn1.Execute (sql)
sql="update 商店 set 店主='"&name&"',id="&id&",說明='"&memo&"',商店類型='"&shoptype&"',注冊資金="&shopmoney&" where  商店名='"&shopname&"'"
conn1.Execute (sql)
set rs1=nothing
conn1.close
set conn1=nothing
%>
</center>
<script language="vbscript">
alert("更改成功！")
window.close()
</script>