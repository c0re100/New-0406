<%'�Ѱ��P��
function jiecmen(fn1,to1,toname)
if toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�Ѱ��P���ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
if info(6)<>"�x��" then
	Response.Write "<script language=JavaScript>{alert('�Ѱ��P���u���x���~�i�H�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �P�� FROM �Τ� WHERE id="&info(9),conn
tongmen1=rs("�P��")
rs.close
rs.open "select �P�� FROM �Τ� WHERE ����='"&fn1&"'",conn
if rs.bof and rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Ӫ������s�b�I');}</script>"
	Response.End

end if
tongmen2=rs("�P��")
if  Instr(1,tongmen1,"|"& fn1 & "|")<>0  then
fn11=Replace(tongmen1,"|"& fn1 &"|","")
fn12=Replace(tongmen2,"|"& info(5) &"|","")
	conn.execute "update �Τ� set �P��='"& fn11 &"' where ����='"&info(5)&"'"
	conn.execute "update �Τ� set �P��='"& fn12 &"' where ����='"&fn1&"'"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Ӫ����å��P�z�������L���I');}</script>"
	Response.End

end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
jiecmen="["&info(5)&"]�P["&fn1&"]�x���N�����M�����Ѱ��P���I"
end function
%>
