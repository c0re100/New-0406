<%
'�ʶR�k�O
function kzdd()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�Ȩ�,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�Ȩ�")<500000000  then
Response.Write "<script language=JavaScript>{alert('�A�����������A�ݭn����5��,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set ���=���+1000000,�Ȩ�=�Ȩ�-500000000,�ާ@�ɶ�=now() where id="&info(9)
kzdd=info(0)& " �h�±z�ʶR�k�O100�U^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>