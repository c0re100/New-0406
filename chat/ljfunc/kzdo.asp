<%
'�ʶR6�ũx��
function kzdo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select grade,�Ȩ�,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�Ȩ�")<3000000000  then
Response.Write "<script language=JavaScript>{alert('�A�����������A�ݭn����30��,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set grade=6,�Ȩ�=�Ȩ�-3000000000,�ާ@�ɶ�=now() where id="&info(9)
kzdo=info(0)& " �h�±z�ʶR6�ũx��^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>