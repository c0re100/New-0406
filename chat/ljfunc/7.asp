<%'�Ѱ��ʸT
function shifang(fn1)
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('�Ѱ��ʸT�ݭn8�ťH�W�޲z���ާ@�I');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���A from �Τ� where �m�W='"& fn1 &"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('�A�ҿ�J���Τ�W���s�b�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("���A")<>"�ʸT" then
	Response.Write "<script language=JavaScript>{alert('["& fn1 &"]�èS���Q�ʸT�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set ���A='���`',�n��=now() where �m�W='" & fn1 & "'"
shifang= info(0) &":<font color=red>�x�����H��"& fn1 &"���ߧ�L�A��{�}�n,�g�j�a�@�ܦP�N���L�@����L�۷s�����|�A�N������I</font>"
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & fn1 & "','�Ѱ��ʸT',now(),'���ߧ�L�A��̥L�o�@���I')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>