<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="jiu.asp"--><%
sql="select * from 帥男 where 姓名='" &info(0)& "'"
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
set connt=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('你弄錯了吧，迎花宮沒有這樣的帥男呀！');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 魅力,銀兩 from 用戶 where id=" & info(9)
set rs=conn.execute(sql)
meili=rs("魅力")
maiyin=meili*200
if rs("銀兩")<maiyin then
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('你的腰包裡沒有"&maiyin&"萬，想來贖身，門都沒！');window.close();</script>"
response.end
else
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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>【還錢贖身】</font><font color=RED>"&info(0)&"</font>花了<font color=FFCFCF>"&maiyin&"兩</font>銀子終於拿回了自己的賣身契！換得清白自由身！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
sql="update 用戶 set 銀兩=銀兩- "&maiyin&"   where id="&info(9)
set rs=conn.execute(sql)
connt.execute("delete * from 帥男 where 姓名='"&info(0)& "'")
connt.close
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('花了"&maiyin&"兩拿了賣身契，歡迎你在沒錢的時候在來我們這工作！');window.close();</script>"
response.end
end if
%>
