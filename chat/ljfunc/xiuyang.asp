<%
'�F�C�׾i
function xiuyang()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �y�O,����,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
Response.Write "<script language=JavaScript>{alert('�A���y�O�����A�ݭn�y�O100,�~��׾i�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set ����=����+10,�y�O=�y�O-100,�ާ@�ɶ�=now() where id="&info(9)
xiuyang=info(0)& "�b���i�|�׾i�A�ݹM�ѤU�_�Ѧ��򤣤֡A���h<font color=red>100</font>�I�y�O,���责��<font color=red>+10</font>�I,�u�O�o���H��!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>