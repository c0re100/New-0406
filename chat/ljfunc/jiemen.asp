<%'����
function jiemen(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�����ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if

if info(2)<5 then
	Response.Write "<script language=JavaScript>{alert('���������u���x���~�i�H�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�P�� FROM �Τ� WHERE �m�W='"&toname&"'" ,conn
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x�����i��P�A�����A�Э��s�ާ@�I');}</script>"
	Response.End
end if
menpai=rs("����")
tongmen=rs("�P��")
if  Instr(1,tongmen,"|"& info(5) & "|")<>0  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�˿��F�a�A���w�g�O�A���P�������F�I');}</script>"
	Response.End
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_tongmen")=toname
jiemen="["&info(5)&"]�x��["&info(0)&"]�V["&menpai&"]�x��{"&toname&"}���X�n�D�����G<input type=button value='�P�N' onClick=javascript:;window.open('jiemenok.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end function
%>
