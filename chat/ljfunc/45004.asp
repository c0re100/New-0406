<%
'�����y
function shuijingqiu(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�����y�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ϥΤ����y�I');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('�x���H�����i�H�i�榹�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("�k�O")<1000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select grade FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs("grade")>=6  then
Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ί����S�ʪ��H���ާ@!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����y' and �ƶq>5 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A5�ӥH�W�������y�ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
rs.close
rs.open "select id FROM �Τ� WHERE id="&info(9),conn
conn.execute "update �Τ� set �k�O=�k�O-1000,�ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set �n��=now()+1/564,���A='�v' where �m�W='"&toname&"'"
conn.execute "update �Τ� set lastkick='"&info(0)&"' where �m�W='"&toname&"'"
shuijingqiu=info(0) & "�q<bgsound src=wav/Ye150.wav loop=1>�f�U�̮��X�����y�A���"&towho&"�f�����������A���������y�_�۪����o�X���~�Ӯg�b<font color=red>"&towho&"</font>�����W�A"&towho&"�g�g�k�k���εۤF." 
call boot(towho)

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>