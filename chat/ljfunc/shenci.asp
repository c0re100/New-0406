<%
'�F�C����
function shenci()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �����[,�D�w,����,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�D�w")<100  then
Response.Write "<script language=JavaScript>{alert('�A���D�w�����A�ݭn�D�w100,�~��h��L�V����I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�����[")<rs("����")*1000  then
conn.execute "update �Τ� set �����[=�����[+5,�D�w=�D�w-100,�ާ@�ɶ�=now() where id="&info(9)
shenci=info(0) & "�b��L�V����,�@�}�J��÷Q��A�𦳩Ү��A���h<font color=red>100</font>�I�D�w,�����W������<font color=red>+5</font>�I,�ǵL��ҡA�`�⦳�Ҧ����!"
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