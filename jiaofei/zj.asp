<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<7  then
	Response.Write "<script Language=Javascript>alert('提示：不是官府中人，不能操作!');window.close();</script>"
	response.end
end if
if Application("ljjh_jiaofei")="土匪" then
	Response.Write "<script Language=Javascript>alert('提示：土匪還在呢！');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_jiaofei")="土匪"
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
sd(199)="<font color=FFD7F4>【消息】突然一群土匪<img src=img/jiao.gif>闖入江湖搶劫，請高手們快去剿匪啊。</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('提示：一群土匪闖入江湖！');window.close();</script>"
response.end
%>
