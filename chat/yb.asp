<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if Application("ljjh_yb")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�p�l�ӱߤF�A���_�S���F...');</script>"
	response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& Application("ljjh_yb") &" where id="&info(9)
yb="���ѯu�O���B�A�X���������{��ڪ����U�W�A���C�C�C�]���G["&name&"]�o��:<img src='img/251.gif'>"&Application("ljjh_yb")&"��!"
Application.Lock
Application("ljjh_yb")=0
Application.UnLock
conn.close
set conn=nothing
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="��"
sd(199)="<font color=FFD060>�i��������j</font>"&yb
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
 