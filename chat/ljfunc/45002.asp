<%
'�e�k�O
function songfali(fn1,to1,toname)
fn1=abs(fn1)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�ǰe�k�O�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ǰe�k�O�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
if Application("ljjh_dubozhang2")=info(0) then
Response.Write "<script language=JavaScript>{alert('�A�{�b�O�k�O�䧽�����a���i�H�ǧO�H�k�O�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�k�O")<fn1 then
Response.Write "<script language=JavaScript>{alert('�A������h���k�O�ܡH�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fn1>50000  then
	if info(4)=0 then 
	Response.Write "<script language=JavaScript>{alert('�D�|���Ǫk�O����W�L5�U,�|������W�L10�U');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	else
	if fn1>100000 then
	Response.Write "<script language=JavaScript>{alert('�D�|���Ǫk�O����W�L5�U,�|������W�L10�U');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
end if
conn.execute "update �Τ� set �k�O=�k�O-"& fn1 &" where id="&info(9)
conn.execute "update �Τ� set �k�O=�k�O+"& fn1 &" where �m�W='"&toname&"'"
songfali=info(0) & "��ۤv��"& fn1 &"�I�k�O�ǰe���F<font color=red>"& toname&"</font>�A"&toname&"�k�O�j�W�I"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function

%>