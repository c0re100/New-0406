<%
sub Gohome
Application("ljjh_active")=0
Application("ljjh_usermdb")="DBQ="+server.mappath("linjiankkk/asz.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
sql="SELECT *  FROM room"
set rs=conn.execute(sql)
 do while Not rs.Eof
Application("ljjh_room")=Application("ljjh_room")&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&";"
	rs.MoveNext
 loop

 rs.close
 set rs=nothing
 set conn=nothing
 Dim nameindex(0)
ljjh_roominfo=split(Application("ljjh_room"),";")
for roomsn=0 to ubound(ljjh_roominfo)-1
	 Application("ljjh_onlinelist"&roomsn)=nameindex
	 fenroom=split(ljjh_roominfo(roomsn),"|")
	 application("ljjh_chatroomname"&roomsn)=fenroom(0)
	Application("ljjh_useronlinename"&roomsn)=""
next 
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
 sd(196)="FFCFCF"
 sd(197)="FFCFCF"
 sd(198)="對"
 sd(199)="<font color=FFCFCF>【系統】</font><font color=red>聊天室開門啦！</font>"
 sd(200)="0"
 Application("ljjh_sd")=sd
 Application("ljjh_line")=1
for roomsn=0 to ubound(ljjh_roominfo)-1
	 fenroom=split(ljjh_roominfo(roomsn),"|")
 Application("ljjh_title"&roomsn)="<font color=FFFFFF style=font-size:9pt>歡迎大家來到靈劍江湖，祝大家聊得開心^_^！</font>"
next
End sub
sub Gouser
 Session.Timeout=10
Application.Lock
Application("ljjh_active")=Application("ljjh_active")+1
Application.UnLock
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