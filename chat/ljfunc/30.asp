<%'�m�Z
function lianwu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���O,��O,�Z�\,����,�Z�\�[,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("��O")<210 or rs("���O")<110 then
	Response.Write "<script language=JavaScript>{alert('�ݭn��O200�A���O100�~�i�H');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�Z�\")<rs("����")*1280+3800+rs("�Z�\�[") then
	conn.execute "update �Τ� set �Z�\=�Z�\+130,��O=��O-200,���O=���O-100,�ާ@�ɶ�=now() where id="&info(9)
	lianwu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\�ߪZ,��O���h-200,���O���h-100,�Z�\����+130,���o��������!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("�Z�\�[")<=rs("����")*1000 then
	conn.execute "update �Τ� set �Z�\�[=�Z�\�[+9,��O=��O-200,���O=���O-100,�ާ@�ɶ�=now() where id="&info(9)
	lianwu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��m�\�ߪZ,��O���h-200,���O���h-100,�Z�\�W������+9,���o��������!"
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