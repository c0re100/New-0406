<%'���� 
function laobao(to1,toname) 
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("ljjh_usermdb") 
rs.open "select ���� from �Τ� where id="&info(9),conn 
grade1=rs("����")
rs.close 
rs.open "select id,�ʧO,�y�O,����,���� from �Τ� where �m�W='"&toname&"'",conn 
idd=rs("id")
meimao=rs("�y�O") 
xingbie=rs("�ʧO") 
grade2=rs("����")
meipai=rs("����")
if grade1<garde2 then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('�H�a����g���A�״I�h�F�A�S��F�A�w�g�����F�I');}</script>" 
Response.End 
end if 
rs.close
rs.open "select ���~�W FROM ���~ WHERE ����='�d��'  and �ƶq>0 and ���~�W='�s�`�d' and �֦���="& idd,conn
if not (rs.eof or rs.bof)  then

conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&idd&" and ���~�W='�s�`�d'")
	Response.Write "<script language=JavaScript>{alert('��設�W�{���s�`�d'�A�A�����ާ@���ѡI');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if meipai="�x��" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('�A�����x���H���i�榹�ާ@�I');}</script>" 
Response.End 
end if 
if xingbie="�k" then 
Response.Write "<script language=javascript>{alert('�A�����k�H�ާ@�I');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
rs.close 
rs.open "select �m�W FROM �W�� WHERE �m�W='"&toname&"'",conn 
if not (rs.eof and rs.bof) then 
Response.Write "<script language=javascript>{alert('�H�a�w�g�O�p�j�F�I');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
rs.close
set rs=nothing
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+500000,¾�~='���B',�D�w=�D�w-1000 where id="&info(9) 
conn.execute "insert into �W��(�m�W,������,����) values ('" & toname & "'," & meimao & ",'"&info(0)&"')" 
laobao=info(0) & "���L�֤l�u�O�n�ΡA�s�X�a�F���N��"& toname &"���ɬ��|�F�A�o��n�B�O50�U�A�D�w�U��1000." 
conn.close
set conn=nothing 
end function 
%>