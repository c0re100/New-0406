<%
'�M��d��
function dogdogdog(fn1,to1,toname)
if toname="�j�a" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('�M������ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W��i�M������I');}</script>"
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
doogdogdog=info(0) & "�n�i���A�A�B��u�O�t�t�r�A���򳣨S���!" 
else
rs.close
conn.execute "update ���~ set ���~�W='�t�¥d' �ƶq=�ƶq+1 where id="& id &" and �֦���=" & info(9)
doogdogdog=info(0) & "�A�b"&towho&"�a�̪��ɩ��U����F<font color=red>�t�¥d</font>�A�H���p�a�}�]�F."
end if

	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>