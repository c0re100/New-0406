<%'�[�J����
function join(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�ʧO FROM �Τ� WHERE id="&info(9),conn
if rs("����")<>"�L" and rs("����")<>"���w" and rs("����")<>"�C�L" then
	Response.Write "<script language=JavaScript>{alert('�����}�A�{�b�������A���a�H');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sex=rs("�ʧO")
rs.close
rs.open "select �A�X FROM ���� WHERE ����='" & fn1 & "' and ����<>'�x��'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('���򤤨S��["& fn1 &"]�o�˪�����.');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if trim(sex)=trim(rs("�A�X")) or trim(rs("�A�X"))="�Ҧ�" then
	conn.execute "update �Τ� set ����='" & fn1 & "', ����='�̤l',grade=1 where id="&info(9)
	conn.execute "update ���� set �̤l��=�̤l��+1 where ����='" & fn1 & "'"
	join="<font color=ffcc00> " & info(0) & " </font>�A���\�a�[�J�F<font color=red>" & fn1 & "</font>�o�Ӫ���!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
Response.Write "<script language=JavaScript>{alert('["&fn1&"]�o�Ӫ������A�X�A�r�I');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>