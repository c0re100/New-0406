<%
function xinqing(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
Response.End 
end if
if len(fn1)>3 and info(2)<10 then
Response.Write "<script language=JavaScript>{alert('�߱����A���פ��i�j��3�Ӧr�I');}</script>"
Response.End
end if

online=Application("ljjh_onlinelist"&session("nowinroom"))
onlinenum=UBound(online)
for i=1 to onlinenum step 6
if online(i+1)=info(0) then

if trim(fn1)<>"���`" then
if trim(fn1)="����" or trim(fn1)="�����`��" or trim(fn1)="�x��" then 
if info(0)<>"�����`��" and  info(0)<>"����" then 
Response.Write "<script language=JavaScript>{alert('�A���O�����A���i�H�p�����I');}</script>"
Response.End
end if
end if
ljjh_zm1=info(4) & "|" & info(5) &"&nbsp;<font size=-1 color=EEEEFF>"&fn1&"</font>" & "|" & info(0) &"|"&info(1)&"|"&info(12)&"|" & info(6)&"|"&info(3)&"|"&info(9)&"|"&info(10)&"|"&info(8)&"|"&info(14)
else
ljjh_zm1=info(4) & "|" & info(5) & "|"&info(0)&"|"&info(1)&"|"&info(12)&"|" & info(6) &"|"&info(3)&"|"&info(9)&"|"&info(10)&"|"&info(8)&"|"&info(14)
end if		
online(i+3)=ljjh_zm1
online(i+4)=sj2
exit for
end if
next
Application.Lock
Application("ljjh_onlinelist"&session("nowinroom"))=online
Application.UnLock
xinqing=info(0)&"�]�m�ۤv�߱����A��["&fn1&"]  "
end function
%>