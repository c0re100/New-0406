<%'�P�N
function dy(to1)
if trim(Application("ljjh_qqq_sf"))<> trim(info(0)) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]�]�S���Q���A�ڡI�I');}</script>"
	Response.End
end if
if trim(Application("ljjh_qqq_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]�]�S���Q���A�ڡI�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���F���@����޲z�B�����v��,�x���T��B�I');}</script>"
	Response.End
end if
    conn.execute "update �Τ� set �Ȩ�=�Ȩ�-5000000,�G�B='"& info(0) &"' where �m�W='"& Application("ljjh_qqq_td") &"'"
    conn.execute "update �Τ� set �Ȩ�=�Ȩ�+5000000,�G�B='"& Application("ljjh_qqq_td") &"' where �m�W='"& info(0) &"'"
if info(1)="�k" then
dy=Application("ljjh_qqq_td") & "�V"& info(0) &"�j�i�A�\,9�k���_,�ֱo"&info(0)&"�s�s�s�n�A�������������F��誺1000�U§���A"&Application("ljjh_qqq_td") & "�ש�����覨���ۤv���p�Ѥ��A�i�ߥi�P�r�I"
else
dy=Application("ljjh_qqq_td") & "�V"& info(0) &"���F�L�Ʋ����e�y�A�åB��F500�U§���A�ש�����覨���ۤv���p�ѱC�A�i�ߥi�P�r�I�p�ѱC�����F�r�u�O�A�֤��L��..."
end if

Application("ljjh_qqq_sf")=""
Application("ljjh_qqq_td")=""
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>