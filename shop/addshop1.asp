<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)<>"江湖總站"  then  response.redirect "../error.asp?id=425"

shopname=Request.Form("shopname")
name=Request.Form("name")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")
shopmoney=Request.Form("shopmoney")

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
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("此商店名已被注冊請重新選擇!")
history.back(1)
</script>
<%Response.End 
end if

sql="insert into 商店 (商店名,店主,id,說明,商店類型,注冊資金) values ('"&shopname&"','"&name&"',"&id&",'"&memo&"','"&shoptype&"',"&shopmoney&")"
conn1.Execute(sql)
set rs1=nothing
conn1.close
set conn1=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【消息】恭喜"&name&"的["&shopname&"]商店開張</font>注冊資金"&shopmoney&"....<font color=FFD7F4>歡迎大家光臨！</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
</center>
<script language="vbscript">
alert("添加成功!")
window.close()
</script>