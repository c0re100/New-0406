<%
function attack(fn1,to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
xx1=rs("�m�W1")
xx2=rs("�m�W2")
rs.close
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select �Z�\,���O,�ŧO,���� FROM �Z�\ WHERE   ����='" & info(5) & "' and �Z�\='"&fn1&"' ",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('�A������:"& menpai &" �èS���o�˪��Z�\["& fn1 &"]�r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
xiuji=rs("�ŧO")
nlin=rs("���O")

if rs("����")="���s" then
rs.close
rs.open "select �ާ@�ɶ�,���A,���O from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=30-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]����,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("���O")<nlin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�����O����["&nlin&"]�L�k�I�i�I');</script>"
	Response.End
end if
if rs("���A")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�w�g���F�r�����H�u���H�I');</script>"
	Response.End
end if
randomize timer
r=int(rnd*120)+200
xiujikill=xiuji*r
nl=xiujikill*4
conn.execute "update �Τ� set ���s=���s+"&xiujikill&",���O=���O-"&xiujikill&"*100,�ާ@�ɶ�=now() where id="&info(9)
attack=xinxi & info(0) & "�I�i["&info(5)&"]��<font color=red>[���s�N]</font><font color=#FFFDAF>"& fn1 &"</font>[<font color=#FFFDAF>"&xiuji&"</font>]��,�ۨ����s�\�ണ��<font color=red>"&xiujikill&"</font>�I,���O����<font color=red>"&nl&"</font>�I!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("����")="��_" then
rs.close
rs.open "select �ާ@�ɶ�,���O from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=30-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]����,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("���O")<nlin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�����O����["&nlin&"]�L�k�I�i�I');</script>"
	Response.End
end if
randomize timer
r=int(rnd*120)+200
xiujikill=xiuji*r
nl=xiujikill*4
conn.execute "update �Τ� set ��O=��O+"&xiujikill&",�ާ@�ɶ�=now() where �m�W='"&toname&"'"
conn.execute "update �Τ� set ���O=���O-"&xiujikill&"*50,�ާ@�ɶ�=now() where id="&info(9)
attack=xinxi & info(0) & "�I�i["&info(5)&"]��<font color=red>[��_�N]</font>"& fn1 &"["&xiuji&"]��,���U<font color=FFFDAF>["&toname&"]</font>������O<font color=red>"&xiujikill&"</font>�I,���O����<font color=red>"&nl&"</font>�I!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("����")="����" then
if toname="" or toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�o�۹ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
xiujikill=xiuji*2
rs.close
rs.open "select ����,�O�@,���A,�Z�\,���s,�ʧO,���O,��O,����,���H��,�ɨ��ɶ�,grade,����,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6 from �Τ� where id="&info(9),conn
grademe=int(rs("grade")*rs("����"))
egj1=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")
efy1=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
if rs("����")="�X�a" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g�X�a�F�A���i�H�A���H�F�I');</script>"
	Response.End
end if
if rs("���A")="��" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g���F�֥h�_���a�I');</script>"
	Response.End
end if
if  rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & info(0) & "]�A���b�O�@���A,�n���[���}�ۤv���O�@�a�I');</script>"
	Response.End
end if
if rs("��O")<300 or rs("���O")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A�{�b�����O����O�p��300�Ф��n�i��o�۩R�O�ާ@�I');</script>"
	Response.End
end if
if rs("���H��")>=30 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�@�c�Ӧh�A�W�L������H�W��30����A�o�ۤF�I');}</script>"	
	Response.End
end if
if rs("���O")<nlin then
	Response.Write "<script language=JavaScript>{alert('�藍�_�A�өۦ��ݭn"&nlin&"�I���O!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
yxjgjfrom=rs("����")
yxjfyfrom=rs("���s")
yxjwgfrom=rs("�Z�\")

xb=rs("�ʧO")
sj=DateDiff("n",rs("�ɨ��ɶ�"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=FFFDAF>�¤O�z��</font>"
	baodoudj=2
end if
wgnl=nlin*xiujikill
nlkill=int(nlin/4)


gw=left(toname,1)
if IsNumeric(gw)=true then
select case info(8)
case "�Ԥh"
yxjgjfrom=yxjgjfrom*1
case "�ɲ��Ԥh"
yxjgjfrom=yxjgjfrom*2
case "���Z�Ԥh"
yxjgjfrom=yxjgjfrom*3
case "���ҾԤh"
yxjgjfrom=yxjgjfrom*4
case "���s�Ԥh"
yxjgjfrom=yxjgjfrom*5
case "�i�h"
yxjfyfrom=yxjfyfrom*1
case "�ʾԫi�h"
yxjfyfrom=yxjfyfrom*2
case "��«i�h"
yxjfyfrom=yxjfyfrom*3
case "���s�i�h"
yxjfyfrom=yxjfyfrom*3
case "����i�h"
yxjfyfrom=yxjfyfrom*4
case "�D�h"
yxjwgfrom=yxjwgfrom*1
case "���D�h"
yxjwgfrom=yxjwgfrom*1.5
case "���k�v"
yxjwgfrom=yxjwgfrom*2
case "���u�H"
yxjwgfrom=yxjwgfrom*2.5
case "���Ѯv"
yxjwgfrom=yxjwgfrom*3
end select
zz=split(toname,"��")
gw=zz(0)
if info(3)>gw*2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('[�H�A�ثe�����ťΤ��ۥ���A�C2�����Ū��Ǫ��F]');</script>"
	Response.End
  end if
if gw>20 and session("nowinroom")=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�n�n���A���n�o�áI');</script>"
	Response.End
end if
rs.close
rs.open "select * from �Ǫ��� where �Ǫ�='"&toname&"'",conn
gwgj=rs("����")
gwfy=rs("���s")
gwtl=rs("��O")
killer=int((yxjgjfrom+wgnl+egj1+yxjwgfrom-gwfy*2)*(baodoudj+xiujikill))
if killer>=180000000 and Application("ljjh_bearcc")=1 then
randomize timer
kaa=int(rnd*99)+1
killer=180000000
killer=killer+kaa
end if
killergw=int(gwgj*2-(yxjwgfrom+yxjfyfrom+efy1))
if killer<=1 then
killer=int(rnd*99)+1
end if
if killergw<=1 then
killergw=int(rnd*99)+1
end if

aa="<img src='../ico/"+gw+".gif'  border='0' width='36' height='36'>"
bb="<img src='../ico/"& info(10) &"-2.gif' width='36' height='36'>"
conn.execute "update �Ǫ��� set kill=kill-"  & killer & " where �Ǫ�='"&toname&"'"
rs.close
rs.open "select * from �Ǫ��� where �Ǫ�='"&toname&"'",conn
if rs("kill")<-100 then
jingye=int((rs("���H��")+1)+gw)
if rs("���H��")>100 then
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('���Ǫ��w�g���`�A�ЧA���W["& ss &"]����,�A�ާ@�I');}</script>"
	Response.End
end if
conn.execute "update �Ǫ��� set ���H��=0 where �Ǫ�='"&toname&"'"
end if
conn.execute "update �Ǫ��� set kill=��O,���H��=���H��+1,�ާ@�ɶ�=now() where �Ǫ�='"&toname&"'"
conn.execute "update �Τ� set allvalue=allvalue+"&jingye&",���O=���O-"&nlin&" where �m�W='"&info(0)&"'"
randomize()
rnd1=int(rnd*info(3))
//rnd1=5
select case rnd1
case 5
attack=xinxi & bb&info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]"& fn1 &"["&xiuji&"]�ŧ���"&aa&"" & towho & ",����" & towho & "�ͩR��:<font color=red>"&killer& "</font>�I�A"& towho &"�G���b�a�W�A["&info(0)&"]��o�԰��g��W�[<font color=FFFDAF>"&jingye&"</font>�I!<img src='../hcjs/jhjs/001/2003.gif' border='0' align='top' alt='"&wupin&"'>"
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�y�P�\',"&info(9)&",'���~',0,1000,0,0,0,1,10000000,2003,0,"&aaao&")"
else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
case else
attack=xinxi & bb&info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]"& fn1 &"["&xiuji&"]�ŧ���"&aa&"" & towho & ",����" & towho & "�ͩR��:<font color=red>"&killer& "</font>�I�A"& towho &"�G���b�a�W�A["&info(0)&"]��o�԰��g��W�[<font color=FFFDAF>"&jingye&"</font>�I!"
end select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
attack=xinxi & info(0) & "�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/26.gif'>����"&aa&""& toname & ",�ϱo"& toname &"��O�U��:<font color=red>"& killer &"</font>�I,"& toname &"�ϧ�r��["&info(0)&"]�ͩR��:<font color=red>"&killergw& "</font>�I!"
conn.execute "update �Τ� set ��O=��O-"  & killergw & ",�ާ@�ɶ�=now() where �m�W='"&info(0)&"'"
rs.close
rs.open "select ��O from �Τ� where �m�W='"&info(0)&"'",conn
if rs("��O")<0 then
conn.execute "update �Τ� set ���A='��',�ާ@�ɶ�=now() where �m�W='"&info(0)&"'"
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=red>�i�t�Ρj</font><font color=FFFDAF>["&info(0)&"]</font>�j�L�Z�ǳy�ڥ����A�G���b"&gw&"�ũǪ���W!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=javascript>alert('�i"&info(0)&"�j�A��������A�G���b�Ǫ���W�F�I');</script>"
	
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if

rs.close
'��誺�P�_
rs.open "select id,�m�W,���A,�O�@,�ʧO,��O,����,���s,�Z�\,����,�|������,����,grade,b1,b2,b3,b4,b5,b6 from �Τ� where �m�W='"&towho&"'",conn
toname=rs("�m�W")
yxjfyto=rs("���s")
yxjwgto=rs("�Z�\")
yxjgjto=rs("����")
jhhy=rs("�|������")
xb1=rs("�ʧO")
menpai1=rs("����")
idd=rs("id")
efy2=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")
gradeto=int(rs("grade")*rs("����"))
if rs("����")<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]�O�s��n�`�N�O�@�z�I');</script>"
	Response.End
end if
if  rs("�O�@")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]���b�O�@���A�A�t�Q��k�a�I');</script>"
	Response.End
end if
if  rs("��O")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]�w�g���F�A�٥��r�I');</script>"
	Response.End
end if
if  rs("grade")>6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]�O�x���޲z�H���A�p��䦳�N���лP�����pô�A���n�H�������I');</script>"
	Response.End
end if
if rs("����")="�X�a" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[" & towho & "]�w�g�X�a�A��������I');</script>"
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
killer=int((yxjgjfrom+wgnl+yxjwgfrom+egj1+grademe-(yxjfyto+yxjwgto+efy2+gardeto))*(baodoudj+xiujikill))

if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update �Τ� set �D�w=�D�w-20,���O=���O-" & abs(nlkill) & ",��O=��O-"& int(killer/10) & " where id="&info(9)
conn.execute "update �Τ� set ��O=��O-"  & killer & " where �m�W='"&towho&"'"
rs.close
rs.open "select ��O from �Τ� where �m�W='"&towho&"'",conn
if rs("��O")>-100 then
if xb="�k" then
	attack=xinxi & info(0) & "�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/26.gif'>����" & toname & ",�ϱo"& toname &"��O�U��:<font color=red>-"& killer &"</font>�I,�ۤv�Q"& toname &"���Z�\�_�o���O�U��:<font color=red>-"& nlkill & "</font>�I,��O:<font color=red>-"& int(killer/4) &"</font>�I."&info(0)&"�԰��g��W�[2�I!"
else
attack=xinxi & info(0) & "�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/26.gif'>����" & toname & ",�ϱo"& toname &"��O�U��:<font color=red>-"& killer &"</font>�I,�ۤv�Q"& toname &"���Z�\�_�o���O�U��:<font color=red>-"& nlkill & "</font>�I,��O:<font color=red>-"& int(killer/4) &"</font>�I."&info(0)&"�԰��g��W�[2�I!"
end if
conn.execute "update �Τ� set allvalue=allvalue+2 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
	conn.execute "update �Τ� set ���H��=���H��+1,allvalue=allvalue+20 where id="&info(9)
	conn.execute "update �Τ� set ���A='��',�ާ@�ɶ�=now(),lastkick='"&info(0)&"' where �m�W='" & towho & "'"
if xx1=info(0) or xx2=info(0) then
conn.execute "update ��D set �m�W1='�L',�m�W2='�L'"
end if
'call boot(toname)
conn.execute "insert into �H�R(����,�ɶ�,����,���]) values ('" & toname & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
rs.open "select id,���~�W,���� from ���~ where (����='�ī~') and �ƶq>0 and �˳�=false and �֦���="&idd,conn
if rs.eof or rs.bof then
if menpai1<>"�q�r��" then 
if xb1="�k" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,"&info(0)&"�԰��g��W�[20�I!"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,"&info(0)&"�԰��g��W�[20�I!"
end if
else
if xb1="�k" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,�]�����q�r�ǡA�x�����y�Ȩ�500�U!"&info(0)&"�԰��g��W�[20�I"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,�]�����q�r�ǡA�x�����y�Ȩ�500�U!"&info(0)&"�԰��g��W�[20�I"
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+5000000 where id="&info(9)
end if
else
id=rs("id")
	wupin=rs("���~�W")
	sm=rs("����")
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
if menpai1<>"�q�r��" then 
if xb1="�k" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,"&info(0)&"�԰��g��W�[20�I!["&toname&"]�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' align='top' border='0' alt='"&wupin&"'></a>"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,"&info(0)&"�԰��g��W�[20�I!["&toname&"]�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' align='top' border='0' alt='"&wupin&"'></a>"
end if
else
if xb1="�k" then
attack=xinxi & info(0) & "<bgsound src=wav/Sol_w2.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,�]�����q�r�ǡA�x�����y�Ȩ�500�U!"&info(0)&"�԰��g��W�[20�I,["&toname&"]�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' align='top' alt='"&wupin&"'></a>"
else
attack=xinxi & info(0) & "<bgsound src=wav/daipu.wav loop=1>�I�i["&info(5)&"]��<font color=red>[�����N]</font>"& fn1 &"["&xiuji&"]��<img src='img/28.gif'>����" & toname & "�A���˹����O"& killer &"�I�A" & toname & "�C�C�ˤU..����q���S�֤F�@��j��,�]�����q�r�ǡA�x�����y�Ȩ�500�U!"&info(0)&"�԰��g��W�[20�I,["&toname&"]�����@�Ǫ��~�A�j�a�ַm!<a href='qiangb.asp?id=" & id&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' align='top'  alt='"&wupin&"'></a>"
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+5000000 where id="&info(9)
end if
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>