<%
'�ʶR10�ŷ|��
function fffff()
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
if rs("�Ȩ�")<10000000000  then
Response.Write "<script language=JavaScript>{alert('�A���Ȩ⤣���A�ݭn�Ȩ�100��,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �|������=10,�Ȩ�=�Ȩ�-10000000000,�ާ@�ɶ�=now() where id="&info(9)
fffff=info(0)& "���\�ʶR10�ŷ|��,�u�O�}�߰�!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>