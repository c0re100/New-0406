<%'�ץ�
function ya(to1,toname,fn1)
if info(2)<4 then
	Response.Write "<script language=JavaScript>{alert('�Q�@����r�A�A���޲z���ťi�����r�I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�I�ץ޹ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(2)<3 then
	Response.Write "<script language=JavaScript>{alert('�A�O���򨭥��H�@���O�x���A�G���O�@�k�ťH�W����I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,���� FROM �Τ� WHERE id="&info(9),conn
grade=rs("grade")
menpai=rs("����")
denji=rs("grade")
rs.close
rs.open "select ����,grade from �Τ� where �m�W='"&toname&"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('�S���o�ӤH�H�A�O���O�ݿ��F�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if denji<6 and menpai<>rs("����") then
	Response.Write "<script language=JavaScript>{alert('�d���F["&toname&"]�]���O�A�̪������r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")<grade  then
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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
application("ljjh_dianxuename")=application("ljjh_dianxuename")&toname&"|"&sj2&"|"&";"
	ya=info(0)& "��" & toname & "�ϥΤF�ץ޳N�A" & toname & "���ʦa���ʤF  [" & fn1 & "]"
else
	ya=info(0) & "��" & toname & "�ϥΤF�ץ޳N�A�i�O�A�����Ť��p�H�a�H�S��k�I"
end if
'�O��
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','�ץ�','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>