<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="homecheck.asp"-->
<!--#include file="checksession.asp"-->
<!--#include file="data.asp"-->
<html>
<%

if session("xiuxi")="1" then
conn.close
%>
<script>
window.alert("一天隻能休息一次！")
history.back()
</script>
<%
else
set ls=conn.execute("select * from 用戶 where 姓名='"&info(0)&"'")
house=ls("house")
select case house
case "house1":
conn.execute("update 用戶 set 體力=體力+30 where id="&info(9))
case "house2":
conn.execute("update 用戶 set 體力=體力+60 where id="&info(9))
case "house3":
conn.execute("update 用戶 set 體力=體力+90 where id="&info(9))
case "house4":
conn.execute("update 用戶 set 體力=體力+120 where id="&info(9))
case "house5":
conn.execute("update 用戶 set 體力=體力+150 where id="&info(9))
case "house6":
conn.execute("update 用戶 set 體力=體力+180 where id="&info(9))
end select
session("xiuxi")="1"
%>
<script>
window.alert("體力已經添加了！")
top.main.location.reload()
history.back()
</script>
<%
conn.close
end if
%>
</html>