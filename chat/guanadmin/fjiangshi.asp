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
alert "���C�C�A���O�x�����H�A�A����񧯩����j�a���I"
window.close()
</script>
<%
response.end
end if
randomize()
jstl=int(rnd*50000)+100
		abc="<a href='kl.asp' target='d'><img src='img/shenmo.GIF' border='0' width='54' height='54'></a>"
		msg="<font color=red>��M�������}�}  �����ǧr�I���@�k�l�y�s�A�@�Y�������i��ѫǡA������O�G+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
Application.Lock
Application("ljjh_kl")=jstl
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="��"
sd(199)="<font color=FFC01F>"&msg&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language="vbscript">
alert "���ߡA�x�����񧯩Ǿާ@���\�I"
window.close()
</script>
<%
%>