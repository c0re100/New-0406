<%'�e��
function give(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�e���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�e���I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� FROM �Τ� WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1>1000000  then
	Response.Write "<script language=JavaScript>{alert('����W�w�e�����i�H�W�L100�U���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�Ȩ�")<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A������h�����ܡH�ݦn�A���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1<100 then
	Response.Write "<script language=JavaScript>{alert('�A����F�a�A����֩�100�⪺�I');}</script>"
	conn.close
	set rs=nothing	
	rs.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & fn1 & " where id="&info(9)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+" & fn1 & " where �m�W='"&toname&"'"
give=info(0) & "��" & fn1 & "��ջȰe���F" & toname &",�o�U�i��"& toname &"�֪�����,�s�n������!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
