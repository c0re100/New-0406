<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="config.asp"-->
<%Response.Expires=0
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
sername=Request.ServerVariables("SERVER_NAME")
ip=Request.ServerVariables("LOCAL_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
if InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")=0 then Response.Redirect "error.asp?id=010"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
disloginname=Application("ljjh_disloginname")
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
userip=Request.ServerVariables("REMOTE_ADDR")
if int(Application("ljjh_chatrs"&session("nowinroom")))>=300 then Response.Redirect "error.asp?id=101"
nickname=Trim(Request.Form("name"))
password=Trim(Request.Form("pass"))
nickname=CStr(Replace(nickname,chr(13)&chr(10),""))
password=CStr(Replace(password,chr(13)&chr(10),""))
gender=CStr(Replace(gender,chr(13)&chr(10),""))
if instr(nickname," ")<>0 then Response.Redirect "error.asp?id=127"
if LCase(nickname)="" then Response.Redirect "error.asp?id=127"
if password="" then Response.Redirect "error.asp?id=128"
if LCase(nickname)=LCase(password) then Response.Redirect "error.asp?id=129"
if server.HTMLEncode(nickname)<>nickname or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 or InStr(nickname," ")<>0 or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 then Response.Redirect "error.asp?id=120"
if server.URLEncode(password)<>password then Response.Redirect "error.asp?id=121"
if server.HTMLEncode(gender)<>gender or InStr(gender," ")<>0 then Response.Redirect "error.asp?id=122"
namelen=0
for i=1 to len(nickname)
zh=mid(nickname,i,1)
zhasc=asc(zh)
if zhasc<0 then
namelen=namelen+2
else
namelen=namelen+1
if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "error.asp?id=120"
end if
next
if namelen>10 then Response.Redirect "error.asp?id=125"
namelen=0
for i=1 to len(gender)
zh=mid(gender,i,1)
zhasc=asc(zh)
if zhasc<0 then
namelen=namelen+2
else
namelen=namelen+1
if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "error.asp?id=122"
end if
next

if namelen>4 then Response.Redirect "error.asp?id=126"
if InStr(LCase(nickname),LCase(disloginname))<>0 or nickname="�j�a" or LCase(nickname)=automan or nickname="��ѫǺ޲z��" or (Instr(nickname,"�_")<>0 or Instr(nickname,"?")<>0) and Instr(nickname,"9")<>0 and nickname<>"�_9�~�h" then Response.Redirect "error.asp?id=130"
if InStr(LCase(nickname),"fuck")<>0 or InStr(LCase(nickname),"sex")<>0 or InStr(nickname,"�l")<>0 or InStr(nickname,"�]")<>0 or InStr(nickname,"�@")<>0 or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"�f")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"�j")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"�Q")<>0 or InStr(nickname,"��")<>0 and InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 or InStr(nickname,"��")<>0 then Response.Redirect "error.asp?id=131"
if InStr(LCase(gender),"fuck")<>0 or InStr(LCase(gender),"sex")<>0 or InStr(gender,"�l")<>0 or InStr(gender,"�]")<>0 or InStr(gender,"�@")<>0 or InStr(gender,"��")<>0 or InStr(gender,"��")<>0 and InStr(gender,"��")<>0 or InStr(gender,"��")<>0 or InStr(gender,"��")<>0 and InStr(gender,"��")<>0 or InStr(gender,"��")<>0 and InStr(gender,"��")<>0 or InStr(gender,"��")<>0 and InStr(gender,"��")<>0 or InStr(gender,"��")<>0 and InStr(gender,"�f")<>0 or InStr(gender,"��")<>0 and InStr(gender,"�j")<>0 or InStr(gender,"��")<>0 and InStr(gender,"�Q")<>0 or InStr(gender,"��")<>0 and InStr(gender,"��")<>0 or InStr(gender,"��")<>0 or InStr(gender,"��")<>0 or InStr(gender,"��")<>0 then Response.Redirect "error.asp?id=132"
dieip=Application("ljjh_dieip")
ipk=split(userip,".",-1)
if Instr(dieip,"*.*.*.*")<>0 or Instr(dieip,ipk(0)&".*.*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&".*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&"."&ipk(2)&".*")<>0 or Instr(dieip,userip)<>0 then Response.Redirect "error.asp?id=111"
iplocktime=50000
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dcz=0
sql="SELECT ip FROM iplocktemp WHERE DateDiff('n',lockdate,#" & sj & "#)>=" & iplocktime
rs.open sql,conn,1,1
if Not(rs.Eof and rs.Bof) then dcz=1
rs.close
if dcz=1 then
sql="DELETE FROM iplocktemp WHERE DateDiff('n',lockdate,#" & sj & "#)>=" & iplocktime
conn.Execute(sql)
end if
sql="SELECT ip,lockdate FROM iplocktemp WHERE ip='" & userip & "'"
rs.open sql,conn,1,1
if NOT(rs.Eof and rs.Bof) then
	lockdate=rs("lockdate")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=110&lockdate=" & server.URLEncode(lockdate)
end if

rs.close
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("ljjh_useronlinename"&i))," "&LCase(nickname)&" ")=0 then ydl=0
	if ydl=1 then 
		Session.Abandon
		Response.Redirect "error.asp?id=140"
		Response.End 
	end if
next 

sql="SELECT �K�X,lastkick FROM �Τ� WHERE �m�W='" & nickname & "'"
rs.open sql,conn
if rs.Eof and rs.Bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=423"
	response.end
end if
regpass=rs("�K�X")
reglastkick=rs("lastkick")
rs.close
temppass=StrReverse(left(password&"godxtll,./",10))
templen=len(password)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
password=replace(mmpassword,"'","B")
'�K�X����
if CStr(password)<>CStr(regpass) then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=141"
end if

if Not(IsNull(reglastkick)) then
	if len(reglastkick)>10 then
		if DateDiff("s",CDate(reglastkick),sj)<=300 then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "error.asp?id=143&lastkick=" & server.URLEncode(reglastkick)
		end if
	end if
end if
if Hour(now())>22 then
rs.open "select * from �����j�� where �D������=''",conn
if rs.eof or rs.bof then
conn.execute "delete * from �����j��"
end if
rs.close
end if
rs.open "SELECT * FROM �Τ� where �m�W='" & nickname & "'",conn,1,3
prevtime=CDate(rs("lasttime"))
value=rs("allvalue")
dengji=int(sqr(value/50))
if DateDiff("m",prevtime,sj)<>0 then
rs("mvalue")=0
end if
rs("����")=dengji
rs("times")=rs("times")+1
rs("lasttime")=sj
rs("lastip")=userip
rs.update
r=Day(date())
s=Hour(time())
f=Minute(time())
sj=r*1440+s*60+f
Session("ljjh_name")=rs("�m�W")
Dim info(16)
info(0)=rs("�m�W")
info(1)=rs("�ʧO")
info(2)=rs("grade")
info(3)=rs("����")
info(4)=rs("�|������")
info(5)=rs("����")
info(6)=rs("����")
info(7)=rs("�v��")
info(8)=rs("¾�~")
info(9)=rs("Id")
info(10)=rs("�W���Y��")
info(11)=rs("�p��")
info(12)=rs("�t��")
info(13)=rs("�ĦW")
info(14)=rs("���H��")
info(15)=rs("lastip")
'if info(2)=11 then
'info(5)="����"
'info(6)="�x��"
'end if
'if info(2)=10 then
'info(5)="�Ư���"
'info(6)="�x��"
'end if
Session("info")=info
erase info
ljjh_hy=rs("�|������")
ljjh_hydate=rs("�|�����")
session("nowinroom")=0
Session.Timeout=20
'��q�r�ǳB�z

if rs("����")="�q�r��" and rs("�D�w")>0 then
conn.execute("update �Τ� set ����='�C�L',����='�̤l' where �m�W='" & nickname & "'")
end if

if rs("�n��")>now() and rs("���A")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420"
	response.end
end if
if rs("���A")="�ʸT" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420"
	response.end
end if
if rs("�n��")>now() and rs("���A")="�c" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=422"
	response.end
end if
if rs("�n��")>now() and rs("���A")="�I��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=480"
	response.end
end if

if int(rs("grade"))>99999999999 and Instr(1,Application("ljjh_admin"),nickname & "|")=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
Response.write "�n�n���A�o�̤��w��«�,�ЧA�X�h,���¦X�@�I" 
response.end
end if

if int(rs("grade"))=999999999999 and Instr(1,Application("ljjh_adminf"),nickname & "|")=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
Response.write "�n�n���A�o�̤��w��«�,�ЧA�X�h,���¦X�@�I" 
response.end
end if
if rs("��O")<-100 or rs("���A")="��" then
xiongshou=rs("lastkick")
conn.close
session.Abandon
Response.Redirect "error.asp?id=421&xiongshou="&xiongshou
response.end
	
end if

if rs("�n��")>now() and rs("���A")="�v" then
xiongshou=rs("lastkick")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=472&xiongshou="&xiongshou
	response.end
end if
dlate=rs("�n��")
if day(dlate)<>day(now()) or month(dlate)<>month(now()) or year(dlate)<>year(now()) then
	conn.execute("update �Τ� set ���H��=0 where �m�W='" & nickname & "'")
end if
conn.execute("update �Τ� set ���A='���`',�O�@=true where �m�W='" & nickname & "'")
wg=rs("�Z�\")
nl=rs("���O")
if rs("�D�w")<-1000 and rs("����")<>"�x��" and rs("����")<>"�x��" then
	conn.execute("update �Τ� set ����='�q�r��',����='�c�H',�O�@=false where �m�W='" & nickname & "'")
end if
if wg<0 then
	conn.execute("update �Τ� set �Z�\=0 where �m�W='" & nickname & "'")
end if
if nl<0 then
	conn.execute("update �Τ� set ���O=0 where �m�W='" & nickname & "'")
end if
if now()-rs("�n��")>0.2 then
	sql="update �Τ� set ��O=��O+50, �Ȩ�=�Ȩ�+1, �n��=now(),���A='���`' where �m�W='" & nickname & "'"
	set Rs=conn.Execute(sql)
end if
if ljjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�@�`�U�Y�i���|������s�hD~`friend��');</script>"
else
	Response.Write "<script Language=Javascript>alert('���ܡG�@�`�U�Y�i���|������s�hD~`friend��');</script>"
end if
if ljjh_hydate<dateadd("d",5,date()) and ljjh_hy>0 then
	hydate=ljjh_hydate-date()
	Response.Write "<script Language=Javascript>alert('���ܡG���s�i,�O�A�ݤ�������F��i��f11');</script>"
end if
if ljjh_hydate<date() and ljjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('���ܡG���ԤH!!!�����D');</script>"
end if

conn.close
set rs=nothing
set conn=nothing
%>
<script>
location.href = "jh.asp"
</script>