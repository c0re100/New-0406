<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(3)<3 then 
Response.Write "<script Language=Javascript>alert('迎花宮不要沒名沒望的江湖小輩！哼！！！');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 性別,魅力 FROM 用戶 WHERE id="& info(9)
set rs=conn.execute(sql)
sex=rs("性別")
meimao=rs("魅力")
maiyin=meimao*100
if sex<>"男" then 
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('暈~你來這裡做什麼？不招小姐哦！！');window.close();</script>"
response.end
end if
if meimao<1000 then 
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('你也太丑了吧~別影響我們生意！去去~！');window.close();</script>"
response.end
end if
%><!--#include file="jiu.asp"--><% 
sql="select * from 帥男 where 姓名='" & info(0) & "'"
set rs=connt.execute(sql)
if rs.bof or rs.eof then
sql="insert into 帥男(姓名,美貌度) values ('" & info(0) & "'," & meimao & ")"
connt.execute sql
conn.execute "update 用戶 set 銀兩=銀兩+"&maiyin&" where id="&info(9)
connt.close
set connt=nothing
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('恭喜，你正式成為迎花宮的帥男，得到賣身銀兩"&maiyin&"兩！等你有錢了再來贖身吧！');window.close();</script>"
response.end
else
connt.close
set connt=nothing
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('你已經是迎花宮的帥男了，還來做什麼呀？');window.close();</script>"
response.end
end if
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
sd(199)="<font color=FFCFCF>【賣身消息】"&info(0)&"</font>為了生計，不得不把自己賣到迎花宮，得到賣身錢<font color=green>"&maiyin&"兩</font>！哎~~何年何月才能還清啊！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>