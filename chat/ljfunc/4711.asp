<%'�s�̤ۧl
function jiamen(fn1,to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
if rs("����")<>"�x��" and rs("����")<>"�ƴx��" and rs("����")<>"����" and  rs("����")<>"�@�k" and rs("����")<>"��D" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�O���򨭥��A�������!');}</script>"
	Response.End
end if
rs.close	
rs.open "select �������,�A�X FROM ���� WHERE ����='" & info(5) & "'",conn
if rs("�������")<15000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A����������Ӥ֤F�A�ܤֱo��1500�U�ڡI');}</script>"
	Response.End
end if
shihe=trim(rs("�A�X"))
rs.close
set rs=nothing
conn.close
set conn=nothing
Application("menpai")=info(5)
jiamen="["&info(5)&"]["&info(0)&"]�V�j�a�s��G["&info(5)&"]�s�̤ۧl,�H�h�O�q�j��,�U��j�L�O�S�ݤF,�֧֥[�J����,�����ۦ�<font color=#009900>"&shihe&"�̤l</font>,���[�J"&Application("menpai")&"����1000�U�Ȩ��  <input type=button style='FONT-SIZE: 9pt'  value='�[�J' onClick=javascript:;disabled=1;window.open('jiamen.asp','d')>"
end function
%>
