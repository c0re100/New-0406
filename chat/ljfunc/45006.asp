<%
'�]�O�p��
function molishi(fn1,to1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o100�I�ڡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-100 where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�]�O�p��' and �ƶq>0 and �֦���=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A���]�O�p�۶ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="& id &" and �֦���=" & info(9)
conn.execute "update �Τ� set �k�O=�k�O+2000 where id="&info(9)
molishi=info(0) & "�q�f�U�̮��X�]�O�p�ۡA�����@��,�]�O�p�۪����{�{�A�ɨ趡<font color=red>"&info(0)&"</font>���k�O�ɺ�2000�I." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>