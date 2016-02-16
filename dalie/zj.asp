<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6   then
	Response.Write "<script Language=Javascript>alert('提示：不是官府中人，不能以操作!');window.close();</script>"
	response.end
end if
if Application("ljjh_dalie")="老虎" then
	Response.Write "<script Language=Javascript>alert('提示：老虎還在呢，你怎麼又放了！');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_dalie")="老虎"
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=FFFDAF>【消息】突然一隻野獸<img src=img/laohu.gif>竄入江湖中傷人，請高手們快去打獵啊。</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('提示：官府放獵物成功！');window.close();</script>"
response.end
%>