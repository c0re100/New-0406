<%'����
function touqian(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�����ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�����A���M�N�|��¼¼���I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,��O,���O,grade from �Τ� where id="&info(9),conn
inyin=rs("�Ȩ�")
if rs("��O")<300 or rs("���O")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A�{�b�����O����O�p��300�Ф��n�i�氽���R�O�ާ@�I');</script>"
	Response.End
end if
rs.close
rs.open "select �m�W,�|������,�Ȩ�,grade from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
jhhy=rs("�|������")
yin=rs("�Ȩ�")
randomize timer
'�|��2�ťH�W�������\�p��5%�I
if jhhy<>0 then
	s=int(rnd*5)+1
else
	s=int(rnd*15)+1
end if
yin=int((yin/100))*s
r=int(rnd*10)

randomize timer
kaa=int(rnd*99)+1
if yin>=1000000 then
randomize timer
yin=1000000
yin=killer+kaa
end if

Select Case r
Case 1
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&yin&" where �m�W='"&toname&"'"
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>�o�F�@�ۭ��s������,����"& toname &"�Ȩ�"& s &"%,�@�p:"& yin &"��!�����p�ߥ����ᥢ�F!"
Case 2
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>�o�F�@�ۭ��s������,��`"& toname &"�P�x���{��,"& info(0) &"�s����§�D�p,�ǷȷȪ����F!"
Case 3
	conn.execute "update �Τ� set ��O=��O-1500 where id="&info(9)
	touqian=info(0) & "<bgsound src=wav/oh_no.wav loop=1>�Q����"& toname &"���Ȩ�,���L�Q�j�a�o�{�F,��O�U�N<font color=red>1500</font>�I!<img src='img/daren.gif'>"
Case 4
	if inyin>50000 then
		conn.execute "update �Τ� set �Ȩ�=�Ȩ�-50000 where id="&info(9)
		touqian=info(0) & "<bgsound src=wav/qt.wav loop=1>��" & toname & "�����Q�o�{,���n�W�U���I��F<font color=red>5�U</font>��,�k�L���T,���B�r!"
	else
		conn.execute "update �Τ� set ���A='��',�Ȩ�=0,�s��=0 where id="&info(9)
		touqian=info(0) & "<bgsound src=wav/oh_no.wav loop=1>�������i" & toname & "�����U,�N�Q�x�����H�o�{�F,�ߧY�Q���h���c�F"
		call boot(info(0))
	end if
case else
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&yin&" where �m�W='"&toname&"'"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&yin&" where id="&info(9)
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>�o�F�@�ۭ��s������,����"& toname &"�Ȩ�"& s &"%,�@�p:"& yin &"��!"& toname &"�j�s,�ڭn�i�A��!"
End Select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>
