<%
function attackd(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('��R�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
info(4)=kkman
if kkman=0 then kkman=1
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�榡�p�U�G[�Z���W&�l�u�ƶq]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
abc=left(trim(zt(1)),1)
if abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�l�u�ƶq�ШϥμƦr�I');}</script>"
Response.End 
end if
aszcc=zt(1)
winname=zt(0)
winname=Replace(winname," ","")
aszcc=Replace(aszcc," ","")
if abs(aszcc)>50 then
	Response.Write "<script language=JavaScript>{alert('�l�u����o�g�W�L50�o�I');}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
xx1=rs("�m�W1")
xx2=rs("�m�W2")
rs.close
rs.open "select ����,�Z�\,�ʧO,���O,����,���H��,�O�@,�l�u,��a�Z��,�Z���¤O,�ɨ��ɶ�,���� from �Τ� where �m�W='"&info(0)&"'",conn
eeee=rs("��a�Z��")
if rs("����")="�X�a" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g�X�a�F�A���i�H�A���H�F�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
dim dwArr

if rs("��a�Z��")<>winname then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�S���t�a"&winname&"�o�تZ���I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("�l�u")<50 then
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�����A�A�S���o��h�l�u�r�H�I�A���l�u�n�ɥR��50�o�~�i�H�g��');}</script>"
Response.End
exit function
end if
if rs("���H��")>=30 then
	conn.execute "update �Τ� set �O�@=false where �m�W='"&info(0)&"'"
	Response.Write "<script language=JavaScript>{alert('�A�@�c�Ӧh�A�W�L������H�W��30����A�o�ۤF�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�O�@")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�����׽m��..����g���I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
okok=rs("�Z���¤O")
if okok=0 then okok=10
wgnl=okok*aszcc
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

'��誺�P�_
rs.open "select �m�W,���A,�ʧO,���s,�Z�\,�|������,�O�@,���� from �Τ� where �m�W='"&toname&"'",conn
toname=rs("�m�W")
yxjfyto=rs("���s")
yxjwgto=rs("�Z�\")
jhhy=rs("�|������")
xb1=rs("�ʧO")
menpai1=rs("����")
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
Response.Write "<script Language=Javascript>alert('["&toname&"]�O�@��..��������I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("���A")="��" then
Response.Write "<script language=JavaScript>{alert('����r;"& toname &"�w�g���F�A�ֵ̽��L�_���r');}</script>"
Response.End
end if
rs.close
randomize timer
kaa=int(rnd*999)+1
killer=int(wgnl*2)
killer=killer*kkman
killer=killer+kaa
if killer>=1000000 then
randomize timer
killer=int(rnd*999999)+1
end if
if killer<=10 then
randomize timer
killer=int(rnd*99)+1
end if
'��誺�P�_
randomize timer
conn.execute "update �Τ� set ��O=��O-"  & killer & " where �m�W='"&toname&"'"
rs.open "select ��O from �Τ� where �m�W='"&toname&"'",conn
if rs("��O")>-100 then
randomize
aaa=int(rnd*6)
select case aaa
case 1
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">��"& winname &"<img src='pic/5.gif'>�g��" & toname & ",�o�g�F<font color=red><b>"&aszcc&"</b></font>�o�l�u,�ϱo"& toname &"��O�U��:<font color=red>-"& killer &"</font>�I."
case 2
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">��"& winname &"<img src='pic/5.gif'>�g��" & toname & ",�o�g�F<font color=red><b>"&aszcc&"</b></font>�o�l�u,�ϱo"& toname &"��O�U��:<font color=red>-"& killer &"</font>�I."
case 3
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">��"& winname &"<img src='pic/5.gif'>�g��" & toname & ",�o�g�F<font color=red><b>"&aszcc&"</b></font>�o�l�u,�ϱo"& toname &"��O�U��:<font color=red>-"& killer &"</font>�I."
case else
attackd=toname&"��<img src='pic/dz44.gif'>"&info(0)&"�j�s�D:����,�u�O���ڷk�o�o,�A�ӧr"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
else
conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+20000,��O=��O+10000,���O=���O+10000 where id="&info(9)
rs.close
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc&">��"& winname &"�g��" & toname & "�uť�@�n�G�s�A"& toname &"���b�a�W���D�G��<img src='pic/dz30.gif'>�ڦ��o�n�G�ڡA�M�Ჴ���@���P�@����F.....<bgsound src=mp3/013.mp3 loop=2>"
set rs=nothing
conn.close
set conn=nothing
end if
end function
%>