<%
function qiangjielin(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('�m�T�O�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
Response.End
exit function
end if
if info(3)<10 then
Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ϥηm�T�O�I');}</script>"
Response.End
end if
if info(2)>6 then
Response.Write "<script language=JavaScript>{alert('�x���H�����i�ާ@�I');}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
if rs("�k�O")<10000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o10000�I�ڡI');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,�s�� FROM �Τ� WHERE �m�W='"&toname&"'",conn
money=int(rs("�s��")/5)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
if rs("����")="�x��"  then
Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ί����S�ʪ��H���ާ@!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�m�T�O' and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A���m�T�O�ܡH');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq-1 where id="& id &" and �֦���="&info(9)
conn.execute "update �Τ� set �k�O=�k�O-10000 where id="&info(9)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&moneyc&" WHERE  id="&info(9)
conn.execute "update �Τ� set �s��=�s��-"&money&" where �m�W='"&toname&"'"
qiangjielin=info(0) & "���X�k��<bgsound src=wav/Phant012.wav loop=1>�m�T�O,���<font color=red>"&towho&"</font>�j�q�@�n:�m�T  ,�q"&towho&"���m�o�s��"&moneyc&"��,"&info(0)&"���Ӫk�O3000�I" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>