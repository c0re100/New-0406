<%'�m��

function qiang(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�m�ܹﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('�x���H�����i�i�榹�ާ@�I');}</script>"
	Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(2)<>1 then
Response.Write "<script Language=Javascript>alert('�n�m�_�Хh�j�U�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �_�� from �Τ� where id="&info(9),conn
if rs("�_��")="�F�C������" then
		Response.Write "<script language=JavaScript>{alert('�A�ۤv���_���٭n�m�H�a���r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
'��誺�P�_
rs.open "select �_��,�_���׽m from �Τ� where �m�W='"&towho&"'",conn

if rs("�_��")="�F�C������"  then
		conn.execute "update �Τ� set �_���׽m=0,�_��='�F�C������' where id="&info(9)
		conn.execute "update �Τ� set �_���׽m=0,�_��='�L' where �m�W='"&towho&"'"
	qiang=info(0) & "��"& towho &"���_��:�F�C�����۷m���A�]�o�즹�_�C�����_���ݭn�i��׽m�~�i�H�o���h���F��I"
		
else
qiang=info(0) & "�Q�m"& towho &"���_��,�i�O"& towho &"���W�èS�������_����!"
		
end if
rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
end function

%>