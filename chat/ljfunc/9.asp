<%'ĵ�i
function jing(fn1,to1)
if info(2)<6 and info(6)<>"�x��" then
	Response.Write "<script language=JavaScript>{alert('ĵ�i�ݭn6�ťH�W�޲z���ާ@�I');}</script>"
	Response.End
end if
if to1="�j�a" or to1=info(0) then
	Response.Write "<script language=JavaScript>{alert('ĵ�i�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
fn1=trim(fn1)
if info(2)<6 and (info(6)="�x��" or info(6)="�ƴx��") then
	jing="<font color=red><b>["&info(5)&"]�̤lť�ۡG" & fn1 & "! </b>(" & info(0) & ")</font>"
else
if info(5)="�x��" then
	jing="<font color=red>�i�x���H���j�Y�¦a��" & to1 & "��:�m�m�m" & fn1 & "�n�n�n����" & info(0) & "</font>"
else
jing="<font color=red>�i�K���ޡj�Y�¦a��" & to1 & "��:�m�m�m" & fn1 & "�n�n�n����" & info(0) & "</font>"
end if
end if
end function
%>