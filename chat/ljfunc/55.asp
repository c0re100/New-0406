<%
function xinqing(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
Response.End 
end if
if len(fn1)>3 and info(2)<10 then
Response.Write "<script language=JavaScript>{alert('心情狀態長度不可大於3個字！');}</script>"
Response.End
end if

online=Application("ljjh_onlinelist"&session("nowinroom"))
onlinenum=UBound(online)
for i=1 to onlinenum step 6
if online(i+1)=info(0) then

if trim(fn1)<>"正常" then
if trim(fn1)="站長" or trim(fn1)="江湖總站" or trim(fn1)="官府" then 
if info(0)<>"江湖總站" and  info(0)<>"熊熊" then 
Response.Write "<script language=JavaScript>{alert('你不是站長，不可以如此做！');}</script>"
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
xinqing=info(0)&"設置自己心情狀態為["&fn1&"]  "
end function
%>