<%'���y
function jiangli(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���y�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<5 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n5�ťH�W�~�i�H���y�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,grade,���� FROM �Τ� WHERE id="&info(9),conn
menpai=rs("����")
if rs("grade")<5 and (rs("����")<>"��D" or rs("����")<>"�@�k" or rs("����")<>"����" or rs("����")<>"�x��" or rs("����")<>"�ƴx��") then
	Response.Write "<script language=JavaScript>{alert('�Q�@����r�A�A�O���򨭥��r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")=5 and (rs("����")="�x��" or rs("����")="�ƴx��") then
	fn1=int(abs(fn1))
if fn1>100000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('���y���i�H�W�L1���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=4 and rs("����")<>"����" then
	fn1=int(abs(fn1))
if fn1>10000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('���y���i�H�W�L1000�U�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=3 and rs("����")<>"�@�k" then
	fn1=int(abs(fn1))
if fn1>1000000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('���y���i�H�W�L100�U�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("grade")=2 and rs("����")<>"��D" then
	fn1=int(abs(fn1))
if fn1>500000 or fn1<1000  then
	Response.Write "<script language=JavaScript>{alert('���y���i�H�W�L50�U�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
rs.close
rs.open "select ������� FROM ���� WHERE ����='" & menpai & "'",conn
if rs("�������")<fn1 then
	Response.Write "<script language=JavaScript>{alert('�����̥u���G"&rs("�������")&"��,���̦�����h�����r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select ���� FROM �Τ� WHERE �m�W='"&toname&"'",conn
if trim(rs("����"))<> trim(menpai) then
	Response.Write "<script language=JavaScript>{alert('�d���F�a�A["&toname&"]�ä��O�A���̤l�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update ���� set �������=�������-" & fn1 & " where ����='" & menpai & "'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+" & fn1 & ",�������=�������-"& fn1 &" where �m�W='"&toname&"'"
jiangli=info(0) & "��ۤv����["&menpai&"]������o���̤l" & toname &","& fn1 & "��@�����y�I"& toname &"�֪�����,�s�n������!"
fn1="�����G"&menpai&"  �Ȩ�G"&fn1
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','���y','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>