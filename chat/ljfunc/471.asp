<%'��D
function diantiao(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('��D�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H��D�I');}</script>"
	Response.End
end if
f=Minute(time())
if f<30 then
	Response.Write "<script language=JavaScript>{alert('�{�b�����D�I��D�ɶ����C�p�ɪ���30����,�Ҧp�G17:30�}�l18:00����!');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W2 FROM ��D",conn
if rs("�m�W2")<>"�L" then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�w�g���H�b��D�F�A�е��@�|�a�Ϊ̽ШD�x���H���Ѱ��O�H��D�I');}</script>"
	Response.End
	end if
rs.close
rs.open "select ����,��O FROM �Τ� WHERE id="&info(9),conn
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x�������i�H��H��D�ڡI');}</script>"
	Response.End
end if
if rs("��O")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A����O�ӧC�F�A�ܤ֤]�o3000�I�H�W�~���H��D�I');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,��O FROM �Τ� WHERE �m�W='"&toname&"'" ,conn
if rs("����")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Z��D�ݭn�A�����Ŧb10�ťH�W�I');}</script>"
	Response.End
end if
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x�������i�H��H��D�ڡI');}</script>"
	Response.End
end if
if rs("��O")<3000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��誺��O�ӧC�F�A�ܤ֤]�o3000�I�H�W�~���H��D�I');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application("dantiao")=toname
diantiao="["&info(0)&"]�V{"&toname&"}���X��D�G"&fn1&"  <input type=button style='FONT-SIZE: 9pt'  value='�����D��' onClick=javascript:;disabled=1;window.open('tiaozhan.asp?id="&info(9)&"&yn=1','d')>"
end function
%>
