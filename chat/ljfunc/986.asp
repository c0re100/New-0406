<%'�W�e�|�O
function give(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�W�e�|�O�ﹳ���F,���i��j�a/�ۤv�ϥ�');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('�԰��ŤӧC,���F����@��,�t�έn5�ũΥH�W�~�i�H�W�e�|�Q�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �|���O FROM �Τ� WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1<1  then
	Response.Write "<script language=JavaScript>{alert('�t�γW�w�̤ӭn��1�ӷ|�O�~�i�H�W�e�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�|���O")<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A�n���S������h���|�O�@�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set �|���O=�|���O-" & fn1 & " where id="&info(9)
conn.execute "update �Τ� set �|���O=�|���O+" & fn1 & " where �m�W='"&toname&"'"
give=info(0) & "��" & fn1 & "�ӷ|�O�e��" & toname &","& toname &"�s�n������!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
