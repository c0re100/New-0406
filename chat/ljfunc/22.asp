<%'��H
function tiren(to1,toname,fn1)
if info(2)<9 then
	Response.Write "<script language=JavaScript>{alert('�޲z�ݭn9�Ū��~�i�H���I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('��H�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs("grade")<info(2) then
	tiren="<bgsound src=wav/gf.wav loop=1>�@�}�g����" & toname & "<IMG SRC=FQ/dz04.gif>��X�F��ѫǡi���=" & info(0) & "�j  ��]�G�m"&fn1&"�n"
	call boot(toname)
else
	Response.Write "<script language=JavaScript>{alert('�]���L�O���ź޲z���A�A�����L�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','��H','"& fn1 &"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>