<%
function jiansu(fn1)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('�޲z�ݭn10�Ū��~�i�H���I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select lastip FROM �Τ� WHERE �m�W='"& fn1 &"'",conn
unlockip=rs("lastip")
conn.Execute "DELETE FROM iplocktemp WHERE ip='" & unlockip & "'"
jiansu="<bgsound src=wav/gf.wav loop=1>" & fn1 & "���ߧ�L�A��������IP!(���=" & info(0) & ")  ["&fn1&"]"

conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& fn1 &"','��H','"& fn1 &"����IP')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>