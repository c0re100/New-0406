<%
function attackc(fn1,to1)
if towho="�j�a" or towho=info(0) then
	Response.Write "<script language=JavaScript>{alert('��R�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
ddkm=info(4)
if ddkm=0 then ddkm=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
xx1=rs("�m�W1")
xx2=rs("�m�W2")
rs.close
rs.open "select ����,�Z�\,�ʧO,���O,����,���H��,�O�@,�ɨ��ɶ�,���� from �Τ� where �m�W='"&info(0)&"'",conn
gcd=rs("����")
if gcd=0 then gcd=1
if rs("����")="�X�a" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g�X�a�F�A���i�H�A���H�F�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("�O�@")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�����׽m��..���ॴ�[�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
yxjgjfrom=rs("����")
yxjwgfrom=rs("�Z�\")
nlfrom=rs("���O")
menpai=rs("����")
xb=rs("�ʧO")
if rs("���H��")>=30 then
	conn.execute "update �Τ� set �O�@=false where �m�W='"&info(0)&"'"
	Response.Write "<script language=JavaScript>{alert('�A�@�c�Ӧh�A�W�L������H�W��"& Application("sjjh_killman") &"����A�o�ۤF�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("�ɨ��ɶ�"),now())
xinxi=""
baodoudj=1
if sj<=60 then
xinxi="<font color=red>�¤O�z��</font>"
baodoudj=3
end if
rs.close
'�}�l�Z�\
rs.open "select �]�k,�򥻧���,���Ӫk�O,����,�I�k����,���� from �ϥΧޯ� where �ޯ�='"&fn1&"' and �m�W='" &info(0)&"'",conn
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A���ޯ�̨èS���o�˪��Z�\["& fn1 &"]�r�I');}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
eepic=rs("����")
mofa=rs("�]�k")
dengji=rs("����")
sm=rs("�I�k����")
wgnl=rs("���Ӫk�O")*dengji
rs.close
'��誺�P�_
rs.open "select �m�W,���A,�ʧO,�O�@,����,�Ȩ�,�s��,���s,�k�O,�k�O�I from �Τ� where �m�W='"&towho&"'",conn
toname=rs("�m�W")
xb1=rs("�ʧO")
aa4=rs("���s")
aa4=int(aa4/10)
aa5=rs("�k�O")
aa6=rs("�k�O�I")
aa7=int(aa5+aa6)
aa7=int(aa7/10)
menpai1=rs("����")
aa1=rs("�Ȩ�")
aa2=rs("�s��")
aa3=int(aa1+aa2)
aa3=int(aa3/100)
if rs("����")="�X�a" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('[" & to1 & "]�w�g�X�a�A��������I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("�O�@")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('["&towho&"]�O�@��..��������I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if xx1=towho or xx2=towho then
if diantiaook=0 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A���i�H�����b��D�����H!');</script>"
Response.End
end if
end if
if rs("���A")="��" then
Response.Write "<script language=JavaScript>{alert('����r;"& towho &"�w�g���F�A�ֵ̽��L�_���r');}</script>"
Response.End
end if
rs.close
randomize timer
kaa=int(rnd*99)+1
killer=int(wgnl*gcd)
killer=killer+kaa
killer=killer-aa4
killer=killer*ddkm
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
killer=180000000
killer=killer+kaa
end if
if killer<=100 then
randomize timer
killer=int(rnd*9999)+1
end if
'��誺�P�_
randomize timer
conn.execute "update �Τ� set ��O=��O-"  & killer & " where �m�W='"&towho&"'"
rs.open "select ID,��O,�|������ from �Τ� where �m�W='"&towho&"'",conn
hyid=rs("ID")
hygd=rs("�|������")
if rs("��O")>-100 then
randomize
aaa=int(rnd*4)
select case aaa
case 1
attackc=xinxi & info(0) & "<bgsound src=3.WAV loop=2>��"& fn1 &"<img src="& eepic &">����" & towho & "(���s�O"&aa4&"),�ϱo"& towho &","&sm&"��O�U��:<font color=red>-"& killer &"</font>.��"& towho &"�i��o�Ȩ�<font color=red><b>"&aa3&"</b>��</font>,�k�O<font color=red><b>"&aa7&"</b>�I</font>"
case 2
attackc=xinxi & info(0) & "<bgsound src=3.WAV loop=2>��"& fn1 &"<img src="& eepic &">����" & towho & ",(���s�O"&aa4&")�ϱo"& towho &","&sm&"��O�U��:<font color=red>-"& killer &"</font>.��"& towho &"�i��o�Ȩ�<font color=red><b>"&aa3&"</b>��</font>,�k�O<font color=red><b>"&aa7&"</b>�I</font>"
case else
attackc=towho&"��<img src='pic/dz44.gif'>"&info(0)&"�j�s�D:����,�u�O���ڷk�o�o,�A�ӧr.��"& towho &"(���s�O"&aa4&")�i��o�Ȩ�<font color=red><b>"&aa3&"</b>��</font>,�k�O<font color=red><b>"&aa7&"</b>�I</font>"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
elseif rs("��O")<-100 and hygd=0 then
conn.execute "update �Τ� set ���A='��',lastkick='"&info(0)&"',�Ȩ�=0,�s��=0,�k�O=0,�k�O�I=0 where �m�W='"&towho&"'"
conn.execute "update �Τ� set ���H��=���H��+2,�Ȩ�=�Ȩ�+"&aa3&",�k�O=�k�O+"&aa7&",��O=��O+10000,���O=���O+10000 where id="&info(9)
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & towho & "',now(),'" & info(0) & "','"&fn1&"')"
conn.execute "delete * from ���~ where �֦���="&hyid
conn.execute "delete * from ������� where �֦���="&hyid
conn.execute "delete * from �ޯ�C�� where �֦���='"&towho&"'"
conn.execute "delete * from �ϥΧޯ� where �m�W='"&towho&"'"
conn.execute "delete * from �歱 where  �֦���='"&towho&"'"
rs.close
attackc=xinxi & info(0) & "<bgsound src=wav/11a.wav loop=3><bgsound src=wav/10.wav loop=3>��"& fn1 &"����" & towho & ","&sm&",���ˤO<font color=red><b>-"& killer &"</b></font>�uť�@�n�G�s�A"& towho &"(���s�O"&aa4&")���b�a�W���D�G��<img src="& eepic &">�ڦ��o�n�G�ڡA�M�Ჴ���@���P�@����F.."&info(0)&"�]����o"& towho &"��d�U��<font color=red><b>"&aa3&"</b>��</font>�Ȥl,�k�O<font color=red><b>"&aa7&"</b>�I</font>"
call boot(toname)
set rs=nothing
conn.close
set conn=nothing
else
conn.execute "update �Τ� set ���A='��',lastkick='"&info(0)&"' where �m�W='"&towho&"'"
conn.execute "update �Τ� set ���H��=���H��+2 where id="&info(9)
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & towho & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
attackc="<bgsound src=wav/11a.wav loop=3><bgsound src=wav/10.wav loop=3>��"& fn1 &"����" & towho & ","&sm&",���ˤO<font color=red><b>-"& killer &"</b></font>�uť�@�n�G�s�A"& towho &"(���s�O"&aa4&")���b�a�W���D�G��<img src="& eepic &">�ڦ��o�n�G�ڡA�M�Ჴ���@���P�@����F.."&info(0)&"�]�������|��,�ҥH����F�賣�S�o��"
call boot(toname)
set rs=nothing
conn.close
set conn=nothing
end if
end function
%>