<%
sub Gohome
Application("ljjh_usermdb")="DBQ="+server.mappath("linjiankkk/asz.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
sql="SELECT �ж��W,����,�����,��F��,�o�ۭ���,�H���ƥ󭭨�,�O�@,�d�� FROM �ж�"
set rs=conn.execute(sql)
conn.close
set conn=nothing
session("nowinroom")=1
 Dim nameindex(0)
 useronlinename=" "
 onliners=0
 Application("ljjh_onlinelist")=nameindex
Application("ljjh_useronlinename")=useronlinename
 Application("ljjh_chatrs")=onliners
 Dim wbq(0)
 Application("ljjh_webicq")=wbq
 webicqname=" "
 Application("ljjh_webicqname")=webicqname
 s=Hour(time())
 f=Minute(time())
 m=Second(time())
 if len(s)=1 then s="0" & s
 if len(f)=1 then f="0" & f
 if len(m)=1 then m="0" & m
 t=s & ":" & f & ":" & m
 Dim sd(200)
 for i=1 to 190
  sd(i)=0
 next
 sd(191)=1
 sd(192)=-1
 sd(193)=0
 sd(194)="AutoMan"
 sd(195)="mingdan"
 sd(196)="FFD7F4"
 sd(197)="FFD7F4"
 sd(198)="��"
 sd(199)="<font color=FFD7F4>�i�t�Ρj��ѫǶ}���աI</font>"
 sd(200)=session("nowinroom")
 Application("ljjh_sd")=sd
 Application("ljjh_line")=1
 Application("ljjh_title")="�w��j�a�Ө��F�C����A���j�a��o�}��^_^�I"

End sub
sub Gouser
 Session.Timeout=10
End Sub

if Application("st_gohome")="" then
call Gohome()
Application("st_gohome")="go"
end if
if Session("Gohome")="" then
call Gouser()
Session("Gohome")="go"
end if
%>