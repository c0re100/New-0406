<%'�e�|��
function give(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�e�|��,���i��j�a/�ۤv�ϥ�');}</script>"
	Response.End
exit function
end if
if info(2)<=10 then%>
	Response.Write "<script language=JavaScript>{alert('�A�޲z�ŨS��11,����e�|�������a�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �|������ FROM �Τ� WHERE id="&info(9),conn
fn1=int(abs(fn1))
end if
conn.execute "update �Τ� set �|������=�|������-0 where id="&info(9)
conn.execute "update �Τ� set �|���O=�|���O+" & fn1 & " where �m�W='"&toname&"'"
give=info(0) & "��" & fn1 & "�ŷ|���e��" & toname &","& toname &"�s�n������!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
