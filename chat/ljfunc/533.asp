<%
function saohuang(to1,toname) 
if toname="�j�a" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('�����ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
Response.End
exit function
end if
if info(1)="�k" then
Response.Write "<script language=JavaScript>{alert('������ʼȥѤk�L����A�k���{�@��I');}</script>"
Response.End
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("ljjh_usermdb") 
rs.open "select �m�W FROM �W�� WHERE ����='"&toname&"'",conn 
if rs.eof and rs.bof then 
Response.Write "<script language=javascript>{alert('���H�S���q�ƹL�����k�I');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-5000000 where �m�W='"&toname&"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000 where �m�W='"&info(0)&"'"
conn.execute "delete * from �W�� where ����='"&toname&"'"
saohuang=info(0) & "�M�@���j�f�̱a�ۤj�M���J�ɬ��|��["&toname&"]��_�Ӭ������~�F�@�y�A�ϥX�F�Q["&toname&"]���i�h���֤k�A["&toname&"]�Q�m��500�U�Ȩ�A["&info(0)&"]�o��100�U." 
rs.close
set rs=nothing
conn.close
set conn=nothing 
end function 
%>