<%'�٭�
function zhanshou(fn1,to1)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('�٭��ݭn10�ťH�W�޲z���ާ@�I');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,grade from �Τ� where id="&to1,conn
toname=rs("�m�W")
if info(2)<rs("grade") then
	Response.Write "<script language=JavaScript>{alert('�A���i�H�ﰪ�ź޲z���ާ@�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set ���A='��',grade=1,����='�C�L',�Z�\=0,���s=0,����=0,�D�w=0,�y�O=0 where id="&to1
call boot(toname)
zhanshou= info(0) &":<font color=red>�x�����H�ۨ�:�H��" & toname & "�٥ߨM!</font>"& fn1
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & toname & "','�٭�',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>