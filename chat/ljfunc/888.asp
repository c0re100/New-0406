<%'���ܨp��
function tracksl()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,�Ѳ� FROM �Τ� WHERE id="&info(9),conn
if rs("�Ѳ�")<>1 then 
tracksl="[�t��]��" & info(0) & "���G�S�o�өR�O�ڡI�A�O���O�d���F�H"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
else
 if session("slshow")=0 then 
  session("slshow")=1 
Response.Write "<script language=JavaScript>{alert('�{�b�A���Ѳ��w�}�q�A�A�i�H�ݨ��誺�Ҧ��ܡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
'tracksl="<font color=red><b>�{�b�A���Ѳ��w�}�q�A�A�i�H�ݨ��誺�Ҧ���!</b></font>" 
  else 
Response.Write "<script language=JavaScript>{alert('�A���Ѳ��w�}�q�F�r�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
'tracksl="<font color=red><b>�A���Ѳ��w�}�q�F�r!</b></font>" 
  end if 
end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
