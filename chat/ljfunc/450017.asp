<%
'�o�g�l�u
function fashezi(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�o�g�l�u�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�o�g�l�u�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ��O,���� FROM �Τ� WHERE �m�W='"&toname&"'",conn
tl=rs("��O")
money=int(rs("��O")/5)
if rs("����")="�x��"  then
Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ί����S�ʪ��H���ާ@!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='����j'  and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A������j�ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�l�u'  and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A���l�u�ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update ���~ set �ƶq=�ƶq-1 where ���~�W='�l�u'  and �֦���="&info(9)
randomize()
rs.close
rs.open "select id FROM �Τ� WHERE id="&info(9),conn
'rnd1=int(rnd()*3)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update �Τ� set ��O=��O-"&money&" where �m�W='"&toname&"'"
fashezi=info(0) & "�q�y�����X�@�����j<bgsound src=wav/Phant012.wav loop=1>�A�c�����a���<font color=red>"&towho&"</font>�N�O�@�j,"&towho&"�ڪ��@�n�A��O���h1/5�A�@�p"&money&"�I......" 
if tl<=-100 then 
conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"
fashezi=info(0) & "�q�y�����X�@�����j<bgsound src=wav/Phant012.wav loop=1>�A�c�����a���<font color=red>"&towho&"</font>�N�O�@�j,"&towho&"�ڪ��@�n�A�ѩ���O����A�Q������šA�n�G��..." 
 call boot(towho)
end if

case else
fashezi=info(0) & "�q�y�����X�@�����j<bgsound src=wav/Phant012.wav loop=1>�A�c�����a���<font color=red>"&towho&"</font>�N�O�@�j,��`�j�k�ӯ�S����......"
end select

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>