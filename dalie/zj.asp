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
	Response.Write "<script Language=Javascript>alert('���ܡG���O�x�����H�A����H�ާ@!');window.close();</script>"
	response.end
end if
if Application("ljjh_dalie")="�Ѫ�" then
	Response.Write "<script Language=Javascript>alert('���ܡG�Ѫ��٦b�O�A�A���S��F�I');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_dalie")="�Ѫ�"
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�����j��M�@�����~<img src=img/laohu.gif>«�J���򤤶ˤH�A�а���̧֥h���y�ڡC</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('���ܡG�x�����y�����\�I');window.close();</script>"
response.end
%>