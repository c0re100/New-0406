<%
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �_��,�_���׽m,����,���O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
if rs("�_��")<>"�F�C������" then
	Response.Write "<script language=JavaScript>{alert('�A�èS�������_��[�F�C������]�ާ@���~�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�_���׽m")>=int(rs("����")/10)+1 then
	conn.execute "update �Τ� set �_��='�L',�k�O=�k�O+100,���O�[=���O�[+�_���׽m,��O�[=��O�[+�_���׽m,�Z�\�[=�Z�\�[+�_���׽m,�D�w�[=�D�w�[+�_���׽m,�y�O�[=�y�O�[+�_���׽m,�����[=�����[+�_���׽m,���s�[=���s�[+�_���׽m,�Ȩ�=�Ȩ�+(200000*�_���׽m) where id="&info(9)
	conn.execute "update �Τ� set �_���׽m=0,�ާ@�ɶ�=now() where id="&info(9)
	xiulian=info(0) & "<bgsound src=wav/xl.wav loop=1>���P�z,�z��A���������_��<font color=red>�F�C������</font>�׽m����,�Ҧ��W���[<font color=red>"& rs("�_���׽m") &"</font>�I,�o�{���G+"& 200000*rs("�_���׽m") &"��,�k�O+100�I,�_���׽m�����A�۰ʮ����F�C�C�C�C�C�C"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<60 then
	s=60-sj
	Response.Write "<script language=JavaScript>{alert('�A�٨S���׽m�����A�е�["& s &"]��A�i��ާ@�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("���O")<150  then
	Response.Write "<script language=JavaScript>{alert('���O�֩�150�A�L�k�׽m�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set �_���׽m=�_���׽m+1,���O=���O-150,�ާ@�ɶ�=now() where id="&info(9)
xiulian=info(0) & "�֦������_���F�C�����۶i��׽m�A�o�O�A��:"& rs("�_���׽m")+1 & "���i���_���׽m�C�C"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>