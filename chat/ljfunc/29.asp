<%'�g��
function jingyan(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('�Ǹg��ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
Response.End
exit function
end if
if info(3)<120 then
Response.Write "<script language=JavaScript>{alert('�Ǹg��ݭn���򵥯�120�Ū��~�i�H�ާ@�I');}</script>"
Response.End
end if
fn1=abs(fn1)
if fn1>500  then
if info(4)=0 then 
Response.Write "<script language=JavaScript>{alert('�D�|���Ǹg��Ȥ���W�L500,�|������W�L10000');}</script>"
Response.End
else
if fn1>10000 then
Response.Write "<script language=JavaScript>{alert('�D�|���Ǹg��Ȥ���W�L500,�|������W�L10000');}</script>"
Response.End
end if
end if
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select allvalue FROM �Τ� WHERE id="&info(9),conn
if rs("allvalue")<fn1 then
Response.Write "<script language=JavaScript>{alert('�A�S������h���g��ȵL�k�ǰe�I');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if

conn.execute "update �Τ� set allvalue=allvalue-" & fn1 & " where id="&info(9)
conn.execute "update �Τ� set allvalue=allvalue+" & fn1 & " where �m�W='"&toname&"'"
jingyan=info(0) & "��" & fn1 & "���g��ǵ��F" & toname &",�Ӧۤv�����򵥯ŭ��C�F�A�j�L�ȧr�I"
'�O��
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','�Ǳ¸g��','"& jingyan & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>