<%
'�ʶR�y�O
function kzdi()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �y�O,�Ȩ�,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�Ȩ�")<700000000  then
Response.Write "<script language=JavaScript>{alert('�A�����������A�ݭn����7��,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �y�O=�y�O+100000000,�Ȩ�=�Ȩ�-700000000,�ާ@�ɶ�=now() where id="&info(9)
kzdi=info(0)& " �h�±z�ʶR�y�O1��^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>