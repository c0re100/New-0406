<%'�Ǳ�
function cuan(fn1,to1,toname)
if toname="" or toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�Ǳ¹ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�Ǥ��O�A�Ǥ��O�ֻ̤ݭn10�šI');}</script>"
	Response.End
end if
fn1=abs(fn1)
if fn1<200 then
Response.Write "<script language=JavaScript>{alert('�ǰe���O�@���̤�200���A�O�Ӥp��I');}</script>"
Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���O FROM �Τ� WHERE id="&info(9),conn
if fn1>30000  then
	Response.Write "<script language=JavaScript>{alert('�Ǥ��O���n�j��3000�n���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>3000 then
	if info(4)=0 then 
	Response.Write "<script language=JavaScript>{alert('�D�|���Ǥ��O����W�L3000,�|������W�L1000000');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	else
		if fn1>1000000 then
		
	Response.Write "<script language=JavaScript>{alert('�D�|���Ǥ��O����W�L3000,�|������W�L1000000');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
		end if
end if
end if
if rs("���O")<fn1 then
	Response.Write "<script language=JavaScript>{alert('�A���̨���h�����O�H�d���F�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set ���O=���O-" & fn1 & " where id="&info(9)
conn.execute "update �Τ� set ���O=���O+" & fn1 & " where �m�W='"&toname&"'"
cuan=info(0) & "�o�\�A���y�q���A�I�I����,�Y�W�C�Ϫ��_�A�ש��" & fn1 & "�����O�ǵ��F" & towho &"�F�A"&  towho  &"�U���P�¡I"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>