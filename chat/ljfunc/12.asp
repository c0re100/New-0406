<%'�l�P�j�k
function xxdf(to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�l���ﹳ�աA�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�l�O�H���O�A���M�N�|��¼¼���I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,���O from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
nlto=rs("���O")
randomize timer
s=int(rnd*20)+1
nlto=int((nlto/100))*s
r=int(rnd*4)
Select Case r
Case 1
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>�ϥΤ@�ۧl�P�j�k,�l��"& toname &"���O<font color=red>"& s &"%</font>,�@�p:<font color=red>"& nlto &"</font>�I!�����p�ߥ����ᥢ�F!"
Case 2
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>�Q�l��"& toname &"���O,�S�����\�I�ۤv���h���O<font color=red>"& s &"%</font>,�@�p:<font color=red>"& nlto &"</font>�I!"& toname &"��,���D�ڼF�`�F�a�I"
	conn.execute "update �Τ� set ���O=���O-"&nlto&" where id="&info(9)
case 3
rs.close
	rs.open "select �Ȩ� from �Τ� where id="&info(9),conn
	if rs("�Ȩ�")>50000 then
		conn.execute "update �Τ� set �Ȩ�=�Ȩ�-50000 where id="&info(9)
		xxdf=info(0) & "<bgsound src=wav/qt.wav loop=1>��" & toname & "�����O�Q�o�{,���n�W�U���I��F<font color=red>5�U</font>��,�k�L���T,���B�r!"
	else
		conn.execute "update �Τ� set ���A='��' where id="&info(9)
		xxdf=info(0) & "<bgsound src=wav/oh_no.wav loop=1>�����f�b" & toname & "�ӻH�W,�N�Q�x�����H�o�{�F,�ߧY�Q���h���c�F"
		call boot(info(0))
	end if
case else
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>�ϥΤ@�ۧl�P�j�k,�l��"& toname &"���O<font color=red>"& s &"%</font>,�@�p:"& nlto &"�I!<font color=red>"& toname &"</font>�I�A��o"& toname &"�R��I"
	conn.execute "update �Τ� set ���O=���O+"&nlto&" where id="&info(9)
	conn.execute "update �Τ� set ���O=���O-"&nlto&" where �m�W='"&toname&"'"
End Select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>
