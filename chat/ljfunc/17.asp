<%'�ذe
function zen(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�ذe�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�榡�p�U�G[���~�W&�ƶq]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
abc=left(trim(zt(1)),1)
if abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�ƶq�ШϥμƦr�I');}</script>"
Response.End 
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�ذe�I');}</script>"
	Response.End
end if
zswupin=zt(0)
wusl=abs(int(left(zt(1),2)))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id,���~�W,����,���O,��O,����,���s,��T��,�Ȩ�,����,����,sm from ���~ where �ƶq>0 and �˳�<>true and ���~�W='" & zswupin & "' and �֦���="& info(9) &" and �ƶq>="&wusl,conn
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A������h�����~�ܡH');}</script>"
	Response.End
end if
id=rs("id")
wupin=rs("���~�W")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
gj=rs("����")
fy=rs("���s")
jgd=rs("��T��")
dj=rs("����")
yin=rs("�Ȩ�")
say=rs("����")
sm=rs("sm")
conn.execute "update ���~ set �ƶq=�ƶq-"&wusl&" where id="&id
rs.close
rs.open "select id from �Τ� where �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id from ���~ where �ƶq>0 and ���~�W='" & zswupin & "' and �֦���=" & idd ,conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~ (���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,����,sm) values ('"&wupin&"','"&idd&"','"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&","&wusl&","&yin&","&dj&",'"&say&"',"&sm&")"
else
	id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+"&wusl&" where id="&id
end if
	zen=info(0) & "��ۤv��[" & zswupin & "]<img src='../hcjs/jhjs/001/"&sm&".gif'>�ذe���F" & toname & "["&wusl&"]�ӡA"&toname&"�ܬO�P�¡I"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>