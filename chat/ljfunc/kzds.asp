<%
'�ʶR10�ũx��
function kzds()
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
if rs("grade")<9  then
Response.Write "<script language=JavaScript>{alert('�Х��ʶR9�ũx��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�Ȩ�")<900000000000000  then
Response.Write "<script language=JavaScript>{alert('�A�����������A�ݭn����9�ʻ�,�~���ʶR�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set grade=10,�Ȩ�=�Ȩ�-8000000000,�ާ@�ɶ�=now() where id="&info(9)
kzds=info(0)& " �h�±z�ʶR10�ũx��^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>