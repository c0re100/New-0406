<%
'�F�C���I
function canchan()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���s�[,�y�O,����,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�y�O")<100  then
Response.Write "<script language=JavaScript>{alert('�A���y�O�����A�ݭn�y�O100,�~��b�s�}�����I�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("���s�[")<rs("����")*1000  then
conn.execute "update �Τ� set ���s�[=���s�[+5,�y�O=�y�O-100,�ާ@�ɶ�=now() where id="&info(9)
canchan=info(0)& "�b�s�}�W�۰��I,�L���}�}�A�𦳩Ү��A���h<font color=red>100</font>�I�y�O,���s�W������<font color=red>+5</font>�I,�ǵL��ҡA�`�⦳�Ҧ����!"
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
set conn=nothing
end function
%>