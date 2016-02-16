<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)>=6 then %>
<!--#include file="const.inc"-->
<%
Response.Buffer=true
youname=info(0)
PK="<font color=#FFff00>PK時間結束，停止PK（單挑也停止）單挑請在PK時間進行，請大家跟聊管合作，違者抓！！！現在開始半小時為賭博時間！！！</font>"
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<br><div align='center'><font color=#ff0000><img src='img/bad.gif' border='0'>"& PK &"<img src='img/bad.gif' border='0'></font></div><br>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%><script language="vbscript">
  alert "執行PK結束操作成功！"
window.close()
</script><%end if%>