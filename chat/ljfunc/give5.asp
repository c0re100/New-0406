<%
'����Z�\
function give5()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �Z�\�[,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('�A�Z�����Z�\���ɶ��٦�["& ss &"]���I�ЦA�����աI');}</script>"
	Response.End
end if
conn.execute "update �Τ� set �Z�\�[=�Z�\�[+90000000,�ާ@�ɶ�=now() where id="&info(9)
give5=info(0)& " �]�s�F�I�t���h�ɭԤF�A�w�g���F�Z�\�W��90000000"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>