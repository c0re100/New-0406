<%'���R
function pingmin(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���R�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<5 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n5�ťH�W�~�i�H���R�I');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('�x���H�����i�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ��O FROM �Τ� WHERE id="&info(9),conn
if rs("��O")<500 then
	Response.Write "<script language=JavaScript>{alert('�A����O�p��500�r�A�S�O��M�H�a���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select grade,��O from �Τ� where �m�W='"&toname&"'",conn
if rs("grade")>=6 then
	Response.Write "<script language=JavaScript>{alert('�A�Q�@ԣ��x�����R�ڡH�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
tl=rs("��O")/2
randomize()
r=int(rnd*6)+1
select case r
case 2
conn.execute "update �Τ� set ��O="&tl&" where �m�W='"&toname&"'"
pingmin=info(0)& "�b���W�j�F�����ĽĦV<bgsound src=wav/7.wav loop=1>["&towho&"]�۹D:�ڸ�A����,�r�r�r,��ť�@�n�z���n�A["&info(0)&"]�������o���H,["&towho&"]������,���K������O�l���@�b�I"
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & info(0) & "',now(),'" & info(0) & "','�P�k���')"
case 3
pingmin=info(0)& "�b���W�j�F�����Ĥj�q�j�s���]��<bgsound src=wav/7.wav loop=1>[" & towho & "]��e���۹D:�ڸ�A���աA�r�r�r�A��ť�@�n�z���n�A["&info(0)&"]�M["&towho&"]�P�k��ɤ@���Q���o���H�I"
conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"

conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & towho & "',now(),'" & info(0) & "','�P�k���')"
case else
pingmin=info(0)& "�b���W�j�F�����Ĥj�q�j�s���]��<bgsound src=wav/7.wav loop=1>["&towho&"]��e���۹D:�ڸ�A���աA�r�r�r�A��ť�@�n�z���n�A["&info(0)&"]�������o���H,["&towho&"]�����{�o�֤~���L�@�T�I"
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & info(0) & "',now(),'" & info(0) & "','�P�k���')"
end select
conn.execute "update �Τ� set ���A='��',��O=0,���O=0,�Z�\=0,����=0,���s=0,�k�O=0 where id="&info(9)
call boot(info(0))
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>