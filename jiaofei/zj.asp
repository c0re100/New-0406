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
	Response.Write "<script Language=Javascript>alert('���ܡG���O�x�����H�A����ާ@!');window.close();</script>"
	response.end
end if
if Application("ljjh_jiaofei")="�g��" then
	Response.Write "<script Language=Javascript>alert('���ܡG�g���٦b�O�I');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_jiaofei")="�g��"
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�����j��M�@�s�g��<img src=img/jiao.gif>���J����m�T�A�а���̧֥h�ϭ�ڡC</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('���ܡG�@�s�g�����J����I');window.close();</script>"
response.end
%>
