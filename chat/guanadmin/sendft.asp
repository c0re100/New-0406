<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)>=7 then
Response.Buffer=true
youname=info(0)
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
onliners=Application("ljjh_chatrs")
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
randomize timer
s=int(rnd*400)+1442
diaox="<font color=#FFC01F>���g���a�A�S����ѫǸ̪��U�j�p���e�W��O��<font color=#FFC01F>"& s &"</font>�I</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
for i=0 to x
	conn.Execute"update �Τ� set ��O=��O+" & s & "  where �m�W='" & online(i) & "'"
next
conn.close
set conn=nothing
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
sd(199)="<font color=#FFC01F>�i�o��O�j"&info(0)&""& diaox &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=javascript>alert('�����I');window.close();</script>"
end if%>
 