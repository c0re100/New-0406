<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6  then
%>
<script language="vbscript">
alert "呵。。你不是官府中人，你不能放怪鳥讓大家打！"
window.close()
</script>
<%
response.end
end if
		
		abc="<a href='dabian.asp' target='d'><img src='img/ying1.gif' border='0' width='94' height='54'></a>"
		msg="<font color=red>一隻怪鳥從冥王星飛臨江湖聊天室，太空鳥啊，一定會有好東西的，大家快打啊!</font><br><marquee width=100%  scrollamount=35>"&abc&"</marquee>"
Application.Lock
Application("ljjh_bianfu")=1
Application.UnLock
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
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="對"
sd(199)="<font color=FFC01F>"&msg&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language="vbscript">
alert "恭喜，官府釋放怪鳥成功！"
window.close()
</script>
<%
%>