<%'�����j��
function bangpai(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�����ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if

if info(2)<5 then
	Response.Write "<script language=JavaScript>{alert('�����j�ԥu���x���~�i�H�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * FROM �Τ� WHERE �m�W='"&toname&"'" ,conn
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x�����i�ާ@�I');}</script>"
	Response.End
end if
menpai=rs("����")
'set rs=conn.execute ("Select count(*) from �Τ� where ����='"&info(5)&"'")
'renshu=rs(0)-1

rs.close
rs.open "select count(*) FROM �Τ� WHERE ����='"&menpai&"' or ����='"&info(5)&"'",conn
renshu=rs(0)-1
if renshu<100 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('������U�̤l�Ӥ֤F�A�ܤֱo��100�������~���}�Ը��I');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM �����j�� WHERE �D������='"&menpai&"'" ,conn
if not rs.bof or not rs.eof then

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�ӽ������ԤF�I');}</script>"
	Response.End
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_bpz")=toname
bangpai="["&info(5)&"]�x��["&info(0)&"]�V["&menpai&"]�x��{"&toname&"}���X�ӽ������j�ԡG<input type=button value='�P�N' onClick=javascript:;window.open('bpzok.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end function
%>
