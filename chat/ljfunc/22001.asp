<%'封鎖IP
function jiafeng(to1,toname,fn1)
if info(0)<>"江湖總站" then
	Response.Write "<script language=JavaScript>{alert('江湖總站才可以使用的！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('封鎖對像錯啦，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select lastip FROM 用戶 WHERE 姓名='"&toname&"'",conn
lockip=rs("lastip")
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
sql="INSERT INTO iplocktemp (ip,lockdate,locker) VALUES ("
sql=sql & SqlStr(lockip) & ","
sql=sql & SqlStr(sj) & ","
sql=sql & SqlStr(toname) & ")"
conn.Execute sql

'conn.Execute "INSERT INTO iplocktemp (ip,lockdate,locker) VALUES ('"& lockip & "','"&SqlStr(sj)&"','"& toname &"')"
rs.close
rs.open "select grade FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs("grade")<info(2) then
	jiafeng="<bgsound src=wav/gf.wav loop=1>" & toname & "因違反江湖規則被封鎖IP30分鐘!(聊管=" & info(0) & ")  ["&fn1&"]"
	call boot(toname)
else
	Response.Write "<script language=JavaScript>{alert('因為他是高級管理員，你不能封他IP！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','踢人','"& fn1 &"封鎖IP')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>