<%'����IP
function jiafeng(to1,toname,fn1)
if info(0)<>"�����`��" then
	Response.Write "<script language=JavaScript>{alert('�����`���~�i�H�ϥΪ��I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('����ﹳ���աA�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select lastip FROM �Τ� WHERE �m�W='"&toname&"'",conn
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
rs.open "select grade FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs("grade")<info(2) then
	jiafeng="<bgsound src=wav/gf.wav loop=1>" & toname & "�]�H�Ϧ���W�h�Q����IP30����!(���=" & info(0) & ")  ["&fn1&"]"
	call boot(toname)
else
	Response.Write "<script language=JavaScript>{alert('�]���L�O���ź޲z���A�A����ʥLIP�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','��H','"& fn1 &"����IP')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>