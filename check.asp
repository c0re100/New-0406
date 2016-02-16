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
if server.HTMLEncode(nickname)<>nickname or InStr(nickname,"▽")<>0 or InStr(nickname,"▼")<>0 or InStr(nickname," ")<>0 or InStr(nickname,"﹛")<>0 or InStr(nickname,"矘")<>0 then Response.Redirect "error.asp?id=120"
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
if InStr(LCase(nickname),LCase(disloginname))<>0 or nickname="大家" or LCase(nickname)=automan or nickname="聊天室管理員" or (Instr(nickname,"稻")<>0 or Instr(nickname,"?")<>0) and Instr(nickname,"9")<>0 and nickname<>"稻9居士" then Response.Redirect "error.asp?id=130"
if InStr(LCase(nickname),"fuck")<>0 or InStr(LCase(nickname),"sex")<>0 or InStr(nickname,"奸")<>0 or InStr(nickname,"淫")<>0 or InStr(nickname,"娼")<>0 or InStr(nickname,"嫖")<>0 or InStr(nickname,"性")<>0 and InStr(nickname,"交")<>0 or InStr(nickname,"妓")<>0 or InStr(nickname,"色")<>0 and InStr(nickname,"黃")<>0 or InStr(nickname,"色")<>0 and InStr(nickname,"情")<>0 or InStr(nickname,"日")<>0 and InStr(nickname,"媽")<>0 or InStr(nickname,"日")<>0 and InStr(nickname,"妹")<>0 or InStr(nickname,"日")<>0 and InStr(nickname,"姐")<>0 or InStr(nickname,"日")<>0 and InStr(nickname,"娘")<>0 or InStr(nickname,"日")<>0 and InStr(nickname,"奶")<>0 or InStr(nickname,"乳")<>0 or InStr(nickname,"陰")<>0 or InStr(nickname,"操")<>0 then Response.Redirect "error.asp?id=131"
if InStr(LCase(gender),"fuck")<>0 or InStr(LCase(gender),"sex")<>0 or InStr(gender,"奸")<>0 or InStr(gender,"淫")<>0 or InStr(gender,"娼")<>0 or InStr(gender,"嫖")<>0 or InStr(gender,"性")<>0 and InStr(gender,"交")<>0 or InStr(gender,"妓")<>0 or InStr(gender,"色")<>0 and InStr(gender,"黃")<>0 or InStr(gender,"色")<>0 and InStr(gender,"情")<>0 or InStr(gender,"日")<>0 and InStr(gender,"媽")<>0 or InStr(gender,"日")<>0 and InStr(gender,"妹")<>0 or InStr(gender,"日")<>0 and InStr(gender,"姐")<>0 or InStr(gender,"日")<>0 and InStr(gender,"娘")<>0 or InStr(gender,"日")<>0 and InStr(gender,"奶")<>0 or InStr(gender,"乳")<>0 or InStr(gender,"陰")<>0 or InStr(gender,"操")<>0 then Response.Redirect "error.asp?id=132"
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

sql="SELECT 密碼,lastkick FROM 用戶 WHERE 姓名='" & nickname & "'"
rs.open sql,conn
if rs.Eof and rs.Bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=423"
	response.end
end if
regpass=rs("密碼")
reglastkick=rs("lastkick")
rs.close
temppass=StrReverse(left(password&"godxtll,./",10))
templen=len(password)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
password=replace(mmpassword,"'","B")
'密碼不對
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
rs.open "select * from 幫派大戰 where 主戰幫派=''",conn
if rs.eof or rs.bof then
conn.execute "delete * from 幫派大戰"
end if
rs.close
end if
rs.open "SELECT * FROM 用戶 where 姓名='" & nickname & "'",conn,1,3
prevtime=CDate(rs("lasttime"))
value=rs("allvalue")
dengji=int(sqr(value/50))
if DateDiff("m",prevtime,sj)<>0 then
rs("mvalue")=0
end if
rs("等級")=dengji
rs("times")=rs("times")+1
rs("lasttime")=sj
rs("lastip")=userip
rs.update
r=Day(date())
s=Hour(time())
f=Minute(time())
sj=r*1440+s*60+f
Session("ljjh_name")=rs("姓名")
Dim info(16)
info(0)=rs("姓名")
info(1)=rs("性別")
info(2)=rs("grade")
info(3)=rs("等級")
info(4)=rs("會員等級")
info(5)=rs("門派")
info(6)=rs("身份")
info(7)=rs("師傅")
info(8)=rs("職業")
info(9)=rs("Id")
info(10)=rs("名單頭像")
info(11)=rs("小孩")
info(12)=rs("配偶")
info(13)=rs("孩名")
info(14)=rs("殺人數")
info(15)=rs("lastip")
'if info(2)=11 then
'info(5)="站長"
'info(6)="官府"
'end if
'if info(2)=10 then
'info(5)="副站長"
'info(6)="官府"
'end if
Session("info")=info
erase info
ljjh_hy=rs("會員等級")
ljjh_hydate=rs("會員日期")
session("nowinroom")=0
Session.Timeout=20
'對通緝犯處理

if rs("門派")="通緝犯" and rs("道德")>0 then
conn.execute("update 用戶 set 門派='遊俠',身份='弟子' where 姓名='" & nickname & "'")
end if

if rs("登錄")>now() and rs("狀態")="獄" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420"
	response.end
end if
if rs("狀態")="監禁" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420"
	response.end
end if
if rs("登錄")>now() and rs("狀態")="牢" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=422"
	response.end
end if
if rs("登錄")>now() and rs("狀態")="點穴" then
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
Response.write "好好玩，這裡不歡迎黑客,請你出去,謝謝合作！" 
response.end
end if

if int(rs("grade"))=999999999999 and Instr(1,Application("ljjh_adminf"),nickname & "|")=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
Response.write "好好玩，這裡不歡迎黑客,請你出去,謝謝合作！" 
response.end
end if
if rs("體力")<-100 or rs("狀態")="死" then
xiongshou=rs("lastkick")
conn.close
session.Abandon
Response.Redirect "error.asp?id=421&xiongshou="&xiongshou
response.end
	
end if

if rs("登錄")>now() and rs("狀態")="眠" then
xiongshou=rs("lastkick")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=472&xiongshou="&xiongshou
	response.end
end if
dlate=rs("登錄")
if day(dlate)<>day(now()) or month(dlate)<>month(now()) or year(dlate)<>year(now()) then
	conn.execute("update 用戶 set 殺人數=0 where 姓名='" & nickname & "'")
end if
conn.execute("update 用戶 set 狀態='正常',保護=true where 姓名='" & nickname & "'")
wg=rs("武功")
nl=rs("內力")
if rs("道德")<-1000 and rs("身份")<>"掌門" and rs("門派")<>"官府" then
	conn.execute("update 用戶 set 門派='通緝犯',身份='惡人',保護=false where 姓名='" & nickname & "'")
end if
if wg<0 then
	conn.execute("update 用戶 set 武功=0 where 姓名='" & nickname & "'")
end if
if nl<0 then
	conn.execute("update 用戶 set 內力=0 where 姓名='" & nickname & "'")
end if
if now()-rs("登錄")>0.2 then
	sql="update 用戶 set 體力=體力+50, 銀兩=銀兩+1, 登錄=now(),狀態='正常' where 姓名='" & nickname & "'"
	set Rs=conn.Execute(sql)
end if
if ljjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：一注冊即可做會員仲唔叫多D~`friend玩');</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：一注冊即可做會員仲唔叫多D~`friend玩');</script>"
end if
if ljjh_hydate<dateadd("d",5,date()) and ljjh_hy>0 then
	hydate=ljjh_hydate-date()
	Response.Write "<script Language=Javascript>alert('提示：當有廣告,令你看不到全部東西可按f11');</script>"
end if
if ljjh_hydate<date() and ljjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：未拉人!!!仲唔快D');</script>"
end if

conn.close
set rs=nothing
set conn=nothing
%>
<script>
location.href = "jh.asp"
</script>