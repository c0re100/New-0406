<%
'�x�����y
function jianglila(fn1,to1,toname)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('�Q�@����r�A�A���޲z���ťi�����r�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� FROM �Τ� WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1>500000 or fn1<1000  then
Response.Write "<script language=JavaScript>{alert('�x�����y���i�H�W�L50�U�֤��i�֩�1000���I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+" & fn1 & " where �m�W='" & toname & "'"
jianglila="�]"& toname &"�樂�򦳰^�m�A"&info(0) & "��x���]�ߪ������o�� " & toname &" "& fn1 & "��@�����y�I"
rs.close
conn.close
set rs=nothing	
set conn=nothing

end function
%>