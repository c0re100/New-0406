<%
function seejj()
if info(5)="�C�L" or info(5)="����" or info(5)="�L" or info(5)="" then
Response.Write "<script language=JavaScript>{alert('�A�٨S�������A����d�ݪ�������I');}</script>"
Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ������� FROM �Τ� WHERE id="&info(9),conn
myjj=rs("�������")
rs.close
rs.open "select ������� FROM ���� WHERE ����='" & info(5) & "'",conn
bpjj=rs("�������")
seejj=info(0) & "�A�{���糧�����^�m�G<font color=red>"&myjj&"</font>��,["&info(5)&"]������Ƭ��G<font color=red>"&bpjj&"��I</font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>