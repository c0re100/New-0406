<%'���
function diu(fn1,to1,toname)
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
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H��󪫫~�I');}</script>"
	Response.End
end if
randomize()
rndok=int(rnd*83876)
zswupin=zt(0)
wusl=abs(int(left(zt(1),2)))
Application.Lock
Application("ljjh_rnd")=rndok
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id,����,sm from ���~ where �ƶq>0 and �˳�<>true  and ���~�W='" & zswupin & "' and �֦���="& info(9) &" and �ƶq>="&wusl,conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('�A������h�����~�ܡH');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
id=rs("id")
lx=rs("����")
sm=rs("sm")
conn.execute "update ���~ set �ƶq=�ƶq-"&wusl&" where id="&id
diu=info(0) & "���ۤv���W�F��Ӧh��ۤv��[" & zswupin & "]<a href='qiangc.asp?id=" & id&"&userint="&rndok&"&db="&wusl&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' alt='"&wupin&"'></a>�ᱼ�F["&wusl&"]�ӡA�֭n�ַm�a�I"
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>