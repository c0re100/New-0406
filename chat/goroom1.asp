<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
chatroomsn=session("nowinroom")
mychatroomsn=lcase(trim(request("roomsn")))
'mychatroomsn=trim(Request.QueryString("roomsn"))
chatroomnamea=Application("ljjh_chatroomname"&mychatroomsn)
chatroomname=trim(Request.QueryString("chatroomname"))
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
if Application("ljjh_chatroomname"&session("nowinroom"))=chatroomnamea then 
response.write "<Script language=javascript>alert('�A�w�g�b["&chatroomnamea&"]���୫�_�i�J�I!');parent.m.location.href='f3.asp';</script>"
  response.end
end if
if mychatroomsn=session("nowinroom") then
Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=000000 background=bk.jpg  bgproperties=fixed>"
  response.write "<Script language=javascript>alert('�A�w�g�b["&chatroomnamea&"]���୫�_�i�J�I!');parent.f3.location.href='f3.asp';</script>"
  response.end
end if
chatroominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(chatroominfo)-1
i=mychatroomsn
online=split(trim(Application("ljjh_useronlinename"&i))," ")
onlinenum=ubound(online)+1
ljchat=split(chatroominfo(i),"|")
num=int(ljchat(1))
if onlinenum>=num  then
  response.write "<Script>alert('["&chatroomnamea&"]�ж���e�H�Ƥw���A����i�J�I');;</script>"
  response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if mychatroomsn=7 then
pkshijian=Hour(now())
if pkshijian=20 then
	rs.open "select * from �����j�� where �D������='"&info(5)&"' or �Ĺ�����='"&info(5)&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
  response.write "<Script>alert('[�H�a�����j�ԧA�M�ͦX����H�I');;</script>"
  response.end
end if
	conn.Execute "update �Τ� set �O�@=false where id="&info(9)
rs.close
else 
  response.write "<Script>alert('[�����j�Ԯɶ�����20�G00-21:00�����I');;</script>"
  response.end
end if
end if
if  info(2)>0 then
	rs.open "select id,���A from �Τ� where id="&info(9)&" and "&ljchat(4),conn
	if rs.eof or rs.bof  then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=000000 background= bk.jpg  bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('�i�J["&ljchat(0)&"]������O�G�G"&ljchat(3)&"');}</script>"
		Response.End
	end if
if rs("���A")="��"  then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=big5'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=000000 background=bk.jpg bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('�A�]�o�ӱߤF�A�w�g��¼¼�F�A�֥h�_���a�I');}</script>"
		Response.End
	end if
	rs.close
end if

ljjh_zm= info(4) & "|" & info(5) & "|"&info(0)&"|"&info(1)&"|"&info(12)&"|" & info(6) &"|"&info(3)&"|"&info(9)&"|"&info(10)&"|"&info(8)&"|"&info(14)
set rs=nothing	
conn.close
set conn=nothing

mychatroomname=Application("ljjh_chatroomname"&chatroomsn)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&chatroomsn)
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubl=UBound(onlinelist)
for i=1 to ubl step 6
 if CStr(onlinelist(i+1))<>CStr(nickname) then
  onliners=onliners+1
  useronlinename=useronlinename & " " & onlinelist(i+1)
  Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
  newonlinelist(js)=onlinelist(i)
  newonlinelist(js+1)=onlinelist(i+1)
  newonlinelist(js+2)=onlinelist(i+2)
  newonlinelist(js+3)=onlinelist(i+3)
  newonlinelist(js+4)=onlinelist(i+4)
  newonlinelist(js+5)=onlinelist(i+5)
  js=js+6
 else
  savetime=onlinelist(i+5)
 end if
next
useronlinename=useronlinename&" "
if onliners=0 then
 dim listnull(0)
 Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
 Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
if Instr(1,Application("ljjh_ying"),"|"& info(0) & "|")=0  then
sd=Application("ljjh_sd")
line=Application("ljjh_line")
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>[���i]"&info(0)&"</font>�I�i�X[���s�T�{]���\�A�಴���K�q[" & mychatroomname & "]�����F�A��ӬO�h�F[" &chatroomnamea& "]!</font>"
	sd(200)=chatroomsn   
	Application("ljjh_sd")=sd
Application.UnLock
end if

session("nowinroom")=mychatroomsn
Application.Lock
onlinelist=Application("ljjh_onlinelist"&mychatroomsn)

dim newonlinelist1()
useronlinename=""
onliners=0
js=1
yjl=0
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if DateDiff("n",onlinelist(i+5),sj)<=60 then
if yjl=0 and len(onlinelist(i+1))>len(info(0)) then

yjl=1
onliners=onliners+2
useronlinename=useronlinename & " " & info(0) & " " & onlinelist(i+1)
Redim Preserve newonlinelist1(js),newonlinelist1(js+1),newonlinelist1(js+2),newonlinelist1(js+3),newonlinelist1(js+4),newonlinelist1(js+5),newonlinelist1(js+6),newonlinelist1(js+7),newonlinelist1(js+8),newonlinelist1(js+9),newonlinelist1(js+10),newonlinelist1(js+11)
newonlinelist1(js)=info(9)
newonlinelist1(js+1)=info(0)
newonlinelist1(js+2)=info(1)&info(4)
newonlinelist1(js+3)=ljjh_zm
newonlinelist1(js+4)=sj
newonlinelist1(js+5)=sj
newonlinelist1(js+6)=onlinelist(i)
newonlinelist1(js+7)=onlinelist(i+1)
newonlinelist1(js+8)=onlinelist(i+2)
newonlinelist1(js+9)=onlinelist(i+3)
newonlinelist1(js+10)=onlinelist(i+4)
newonlinelist1(js+11)=onlinelist(i+5)
js=js+12

else
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist1(js),newonlinelist1(js+1),newonlinelist1(js+2),newonlinelist1(js+3),newonlinelist1(js+4),newonlinelist1(js+5)
newonlinelist1(js)=onlinelist(i)
newonlinelist1(js+1)=onlinelist(i+1)
newonlinelist1(js+2)=onlinelist(i+2)
newonlinelist1(js+3)=onlinelist(i+3)
newonlinelist1(js+4)=onlinelist(i+4)
newonlinelist1(js+5)=onlinelist(i+5)
js=js+6
end if
end if
next

if yjl=0 then
onliners=onliners+1
useronlinename=useronlinename & " " & info(0)
Redim Preserve newonlinelist1(js),newonlinelist1(js+1),newonlinelist1(js+2),newonlinelist1(js+3),newonlinelist1(js+4),newonlinelist1(js+5)
newonlinelist1(js)=info(9)
newonlinelist1(js+1)=info(0)
newonlinelist1(js+2)=info(1)&info(4)
newonlinelist1(js+3)=ljjh_zm
newonlinelist1(js+4)=sj
newonlinelist1(js+5)=sj
end if

useronlinename=useronlinename&" "
if onliners=0 then
dim listnull1(0)
Application("ljjh_onlinelist"&mychatroomsn)=listnull1
else
Application("ljjh_onlinelist"&mychatroomsn)=newonlinelist1
end if

Application("ljjh_useronlinename"&mychatroomsn)=useronlinename
Application("ljjh_chatrs"&mychatroomsn)=onliners

session("nowinroom")=mychatroomsn

%>
<script>
top.location.reload();
</script>