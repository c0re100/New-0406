<%'���ؾi��
function bimu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���O,��O,����,��O�[,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("���O")<110  then
	Response.Write "<script language=JavaScript>{alert('�ݭn��O100�~�i�H�ާ@�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("��O")<rs("����")*240+5260+rs("��O�[") then
	conn.execute "update �Τ� set ��O=��O+1500,���O=���O-100,�ާ@�ɶ�=now() where id="&info(9)
	bimu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\����,���O���h<font color=red>-100</font>�I,��O����<font color=red>+1500</font>�I,���o��������!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("��O�[")<=rs("����")*1600 then
	conn.execute "update �Τ� set ��O�[=��O�[+9,���O=���O-100,�ާ@�ɶ�=now() where id="&info(9)
	bimu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\����,���O���h-100��O�W������+9,���o��������!"
else
	Response.Write "<script language=JavaScript>{alert('�{�b�A���W�����F�A���ɤF�ŦA�m�a');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
