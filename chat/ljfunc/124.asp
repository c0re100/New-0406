<%
'�ʶR6�ŷ|��
function bbbbb()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,�|������,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=1-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("�Ȩ�")<6000000000  then
Response.Write "<script language=JavaScript>{alert('�A���Ȩ⤣���A�ݭn�Ȩ�60��,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �|������=6,�Ȩ�=�Ȩ�-6000000000,�ާ@�ɶ�=now() where id="&info(9)
bbbbb=info(0)& "���\�ʶR6�ŷ|��,�u�O�}�߰�!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>