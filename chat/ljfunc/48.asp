<%'���B����
function xingyu(fn1)
if Application("ljjh_xinyu")=0 or isempty(Application("ljjh_xinyu")) then
	randomize()
	rnd1=cint(rnd()*999)+1
	Application.Lock
	Application("ljjh_xinyu")=rnd1
	Application.UnLock
end if
fn1=int(abs(fn1))
if fn1<0 or fn1>1000 then
	Response.Write "<script language=JavaScript>{alert('��J���~,���B�������G1-1000�������ơI');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if fn1=Application("ljjh_xinyu") then
	Application.Lock
	Application("ljjh_xinyu")=0
	Application.UnLock
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+5000000 where id="&info(9)
	
	xingyu="���F�C�m�����G����"&info(0)&"�z���F�F�C����֧Q�����A���X�G"&fn1&"<img src='img/251.gif'>����500�U<img src='img/251.gif'>�A�j�a��ܮ��ߡI"
else

	rs.open "select �Ȩ� FROM �Τ� WHERE id="&info(9),conn
	if rs("�Ȩ�")<1000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�S�����F�A�S��k�F�I�O�R�F�I');}</script>"
		Response.End
	end if
	
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-2000 where id="&info(9)
	
	xingyu="���F�C�m�����G"&info(0)&"�ʶR�F�����A���F�C�֧Q�Ʒ~�@�^�m�I���X�G"&fn1&"�S������,�٦��U�@���A�٦����|�I�����G500�U"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>
