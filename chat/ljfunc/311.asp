<%'��D����
function quxiaoa()
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('�A���O�x�����H�A�Ф��n�ާ@�I');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W1,�m�W2 FROM ��D",conn
xx1=rs("�m�W1")
xx2=rs("�m�W2")
Application.Lock
Application("diantiao")=""
Application.unLock
'conn.execute "delete * from ��D where �m�W1<>'�L'"
conn.execute "update ��D set �m�W1='�L',�m�W2='�L'"
quxiaoa="�x���H��"&info(0)& "�]�ѻP�D�Ԫ�����["&xx1&"]----["&xx2&"]�]���Y�ǭ�]�����D�ԡA��L�H�i�H�i��D�ԤF  "

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>