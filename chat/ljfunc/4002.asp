<%
'�M���s�]
function dddddd(fn1,to1,toname)
if toname="�j�a" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('�M�s�]�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W��i�M�y�P�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=10-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]����,�A�ާ@�I');}</script>"
	Response.End
end if
if fla<10000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o10000�I�ڡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-100000,�ާ@�ɶ�=now() where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
dddddd=info(0) & "�n�i���A�A�M�M�F����U�Ө����]�S����줰���s�],"&info(0) & "�l��100000�I�k�O!" 
else
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�s�]' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����) values ('�s�]',"&info(9)&",'���~',100,0,0,0,0,1,200000,1600)"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1 where id="& id &" and �֦���=" & info(9)
	end if
dddddd=info(0)& "�A�d���U�W�M�X<font color=red>�s�]</font>���U���A�o���Ʀb�@�������̳Q�A���F�A����~�~�s�]���������]�O�a." 
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>