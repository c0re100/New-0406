<%'����
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ��O,����,���O�[,���O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("��O")<210  then
	Response.Write "<script language=JavaScript>{alert('��200�H�W����O�~�i�H�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("���O")<rs("����")*64+2000+rs("���O�[") then
	conn.execute "update �Τ� set ���O=���O+30,��O=��O-200,�ާ@�ɶ�=now() where id="&info(9)
	dazhuo=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\����,��O���h-200���O����+30,���o��������!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("���O�[")<=rs("����")*800 then
	conn.execute "update �Τ� set ���O�[=���O�[+9,��O=��O-200,�ާ@�ɶ�=now() where id="&info(9)
	dazhuo=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\����,��O���h<font color=red>-200</font>�I,���O�W������<font color=red>+9</font>�I,���o��������!"
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