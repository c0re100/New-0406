<%'���{
function stu(to1)
if to1="�j�a" or to1=info(0) then
	Response.Write "<script language=JavaScript>{alert('���{�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H���{�I');}</script>"
	Response.End
end if
if trim(Application("ljjh_bais_sf"))<> trim(info(0)) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]�]�S���Q���A���v�I');}</script>"
	Response.End
end if
if trim(Application("ljjh_bais_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]�]�S���Q���A���v�I');}</script>"
	Response.End
end if
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-50000,�v��='"& info(0) &"',�v�ť��='0' where �m�W='"& Application("ljjh_bais_td") &"'"
stu=Application("ljjh_bais_td") & "�V"& info(0) &"��ǤF5�U�����v�O�A�S�O�I�Y�S�O���y���A�ש�D�o"& info(0) &"���ۤv���{,���n�v�Ŧb�A�Z�\�j�i���I"
Application("ljjh_bais_sf")=""
Application("ljjh_bais_td")=""
conn.close
set conn=nothing
end function
%>