<%
'�e��
function songkuan(fn1,to1,toname)
fn1=abs(fn1)
if fn1<10000 or fn1>50000000 then
	Response.Write "<script language=JavaScript>{alert('���̤�1�U�A�̦h5000�U�I');}</script>"
	Response.End
end if
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('�A�L�v�ާ@�I');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �s�� from �Τ� where id="&info(9),conn
conn.execute "update �Τ� set �s��=�s��+" & fn1 & " where �m�W='"&toname&"'"
songkuan=info(0) & "���F�C����Ȳ�<img src='pic/Dz01.gif'>:"& fn1 &"��e��"& toname &"�H�����y!"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>