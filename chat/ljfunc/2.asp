<%'�I��
function dian(to1,fn1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�I�޹ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('�Q�@����r�A�A�i���O�x�����I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,�m�W from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�S���o�ӤH�H�A�O���O�ݿ��F�I');}</script>"
	
	Response.End
end if
if rs("grade")<info(2)  then
	call boot(toname)
	conn.execute "update �Τ� set �n��=now()+0.02,���A='�I��' where �m�W='"&toname&"'"
	dian=info(0)& "��" & toname & "�ϥΤF�I�޳N�A" & toname & "�b�b�a���ʤF  ��28�����~�i�H�W�u  "
else
	dian=info(0) & "��" & toname & "�ϥΤF�I�޳N�A�i�O�L�O�x�������ź޲z���I"
end if
'�O��
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','�I��','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
