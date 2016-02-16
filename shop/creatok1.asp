<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
shopname=Request.Form("shopname")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")

if shopname="" or memo="" or shoptype="請選擇" then 
%><script language="vbscript">
alert("每項必填請返回重填!")
history.back(1)
</script>
<%Response.End 
end if
%><!--#include file="data.asp"--><%
rs.open "select 會員等級,等級,銀兩 FROM 用戶 WHERE id="&info(9),conn
if rs("會員等級")=0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，商店老板目前隻對會員開放！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("銀兩")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，你需要帶齊2億銀兩作為商店的注冊資金！');location.href = 'window.close()';</script>"
	Response.End
end if
if rs("等級")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，開設商店需要你的戰鬥級在100級以上！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE (類型='鮮花' or 類型='右手' or 類型='左手' or 類型='藥品' or 類型='裝飾' or 類型='盔甲') and  數量>0 and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，開設商店需要你有6種以上物品！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
conn.execute "update 用戶 set 銀兩=銀兩-200000000 where id="&info(9)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
sql="SELECT * FROM 商店 where 商店名='"&shopname&"' or 店主='"&info(0)&"'"
set rs1=conn1.Execute (sql)
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("此商店名已被注冊或你已經是老板了請重新選擇!")
window.close()
</script>
<%Response.End 
end if

sql="insert into 商店 (商店名,店主,id,說明,商店類型,注冊資金) values ('"&shopname&"','"&info(0)&"',"&info(9)&",'"&memo&"','"&shoptype&"',200000000)"
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
sd(199)="<font color=FFD7F4>【消息】恭喜"&info(0)&"的["&shopname&"]商店開張</font>注冊資金2億....<font color=FFD7F4>歡迎大家光臨！</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
</center>
<script language="vbscript">
alert("申請成功!")
window.close()
</script>