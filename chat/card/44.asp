<%@ LANGUAGE=VBScript codepage ="950" %>

<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####�ж��B�z#####
zhbg_roominfo=split(Application("zhbg_room"),";")
chatinfo=split(zhbg_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('���ܡG�b["&chatinfo(0)&"]�ж����i�H�ϥΥd���I');}</script>"
	Response.End
end if
if Instr(LCase(Application("zhbg_useronlinename"&nowinroom)),LCase(" "&zhbg_name&" "))=0 then 
	Session("zhbg_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
f=Minute(time())
zhbg_kptime=Application("zhbg_kptime")
if f>zhbg_kptime and chatinfo(0)<>"���象��" then
 Response.Write "<script language=javascript>{alert('���ܡG�D����ж��Υd�Цb�C�p�ɪ��e"&zhbg_kptime&"��');}</script>"
 Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'������}�B�I�ץާP�_
'call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
'��t�θT��r�ųB�z
if zhbg_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
if fnn1<>"�k�ӥd" then call dianzan(towho) 
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�d��
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
if zhbg_name=to1 and instr(";�Ѱ��d;���`�d;�d�|�d;�Y���d;�o�åd;�e���d;��H�d;�V�v�d;�m�˥d;���H�d;�d�|�d;�a���d;�I���d;�����d;����d;���ťd;���~�d;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('�i"&fn1&"�j����ۤv�i��ާ@�I');</script>"
	Response.End
end if
if to1="�j�a" and instr("�e���d;��H�d;�V�v�d;�m�˥d;���H�d;�d�|�d;�a���d;�I���d;�����d;����d;���ťd;���~�d;�ؤl�d;�ت�d;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('�i"&fn1&"�j����j�a�i��ާ@�I');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("zhbg_usermdb")
fn1=trim(fn1)
rs.open "select ��� from �Τ� where  �m�W='"&to1&"'",conn,2,2
zstt=rs("���")
if zstt>=8 then
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�Q��%%�ϥ�["&fn1&"]�A��%%�w�g�m���F�W�Z����O(��ͤj�󵥩�8��)�A���d���w�g�Q�ƸѤF...</font>"
exit function
end if
rs.close
rs.open "select �|�����d,w5,���� from �Τ� where  �m�W='"&zhbg_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "�֯��d"
	conn.Execute ("update �Τ� set sl='�֯�',slsj=now()+3 where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b���F�ᨭ�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�����ϥΤF["&fn1&"]<img src=card/BXK.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA�z�P�ڦP�b...</font>"
case "�]���d"
	conn.Execute ("update �Τ� set sl='�]��',slsj=now()+3 where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b���F�ᨭ�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�����ϥΤF["&fn1&"]<img src=card/caishen.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA�z�P�ڦP�b...</font>"
case "�a���d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set sl='�a��',slsj=now()+3 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b���F�P["&to1&"]�ᨭ�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]<img src=card/qiongshen.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA�P�A�P�b...</font>"
case "�I���d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set sl='�I��',slsj=now()+3 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b���F�P["&to1&"]�ᨭ�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]<img src=card/aishen.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA�P�A�P�b...</font>"
case "�����d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set sl='����',slsj=now()+3 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b���F�P["&to1&"]�ᨭ�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]<img src=card/sishen.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA�P�A�P�b...</font>"
case "�e���d"
	conn.Execute ("update �Τ� set sl='�L',slsj=now() where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]���\�Ѱ��F���F�ᨭ�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF["&fn1&"]<img src=card/BXK.JPG><bgsound src=mav/XL.MID loop=1>�A���ڡA88�Isee tomorrow!</font>"
case "���]�d"
	conn.Execute ("update �Τ� set ���B����=���B����+5  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('�z�ϥΤF���]�d�A�]�N��ܧA���F�u�������]�C�I');}</script>"
	kapian="<font color=red>�i���]�d�j<font color=" & saycolor & ">##�ϥΤF���]�d<img src=card/semo.JPG><bgsound src=mav/XL.MID loop=1>�A�����o�U�A���C���˰��A���}�s�����R�H���n�A�F�A���B���Ʀ�5���F�G�����A���]�S��k�F�A�����O�ӦⰭ�O�A�����C�j�a�֥��o�Ӥj���]�A�֬ݬݧr�A�~�X�ѡA�w�g���L�o��h���B�F�C�����C</font>"
case "�D§�d"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]<img src=card/feili.JPG><bgsound src=mav/XL.MID loop=1>...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �Τ� set �y�O=�y�O-500 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]���y�O�֤F500�I');}</script>"
	kapian="<font color=red>�i�d���j<font color=" & saycolor & ">##��}�}���A�w�I�I���N�V%%�����p�L�@�}KISS,%%�uı�@�}���ߡC�y�O�֤F500�I..</font>"
case "�X�n�d"
	conn.Execute ("update �Τ� set ����=����+10 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�ͽˤ���,�u�۪��I�ꩼ��,���ڭּ̧֡A���ڭ̦X�n.�I');}</script>"
	kapian="<img src=card/he1.gif><bgsound src=wav/hua.wav loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF�X�n�d<img src=card/hehao.gif>,����e�W�F10�Ӫ���.�u�O����,�B�ͽФ��n�@�Ǥp��,�N���h�F�ڭ̪��ͽ�.�@�ڭ̭��k��n,�ͽ˦a�[�Ѫ�,�@�_��,�B�ͤ@�ͤ@�_��,���Ǥ�l���b��,�@�y��,�@���l,�@�ͱ�,�@�M�s...</font>"
case "���C�d"
	conn.Execute ("update �Τ� set ����=����+1000 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]����׳̰��̬��̱j�����t�F���C,������m�W�F1000����,�u�O�p�N�_��!');}</script>"
	kapian="<img src=card/jian.gif><bgsound src=readonly/okok.wav loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF���ʽ��̱j�̦n�̬����d-------���C�d<img src=card/jian1.gif>,�ڤۤ��C�b�����o�eģ�������m.����e�W�F1000����,�o�^�ѤU�L�ĤF,%%�u�O�ӷP�ʤF��.�����A�i���i�H�b�Ӥ@�i�r:)...</font>"
case "�]�ޥd"
	conn.Execute ("update �Τ� set ���m=���m+1000 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�D�Z����`���̰��̬��̱j�������ʸt�s����,������m�W�F10000���m!');}</script>"
	kapian="<img src=card/dun.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF�D�Z����`���̱j�̦n�̬����d-------�]�ޥd<img src=card/dun1.gif>,�ڤۤ��ަb�����o�eģ�������m.����e�W�F1000���m,�o�^�ѤU�L�ĤF,%%�u�O�ӷP�ʤF��.�����A�ϥ@�D�r:)...</font>"
case "�ѧ٥d"
    conn.Execute ("update �Τ� set �k�O=�k�O+20000 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�D�Z����`���̰��̬��̱j�����t�s����,������m�W�F20000�k�O�I');}</script>"
	kapian="<img src=card/jie.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF�D�Z����`���̱j�̦n�̬����d-------�ѧ٥d<img src=card/jie1.gif>,�ڤۤ��٦b�����o�eģ�������m.����e�W�F2000�k�O,�o�^�ѤU�L�ĤF,%%�u�O�ӷP�ʤF��.�����A�ϥ@�D�r:)...</font>"
case "�ԯ��d"
    conn.Execute ("update �Τ� set �Z�\=�Z�\+2000 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�D�Z��׺��̰��̬��̱j�����t�����b,������m�W�F2000�Z�\�I');}</script>"
	kapian="<img src=card/zhan.gif><bgsound src=cd.mid loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF�D�Z����`���̱j�̦n�̬����d-------�ԯ��d<img src=card/zhan1.gif>,�ڤۤ��b�b�����o�eģ�������m.����e�W�F2000�Z�\,�o�^�ѤU�L�ĤF,%%�u�O�ӷP�ʤF��.�����A�ϥ@�D�r:)...</font>"
case "�P�u�d"
    conn.Execute ("update �Τ� set allvalue=allvalue+3000 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�D�Z����`���̰��̬��̱j�����t�P���u,������m�W�F30000�s�I�I');}</script>"
	kapian="<img src=card/xue.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##�u�۱o��%%�ϥΤF�D�Z��׺��̱j�̦n�̬����d-------�P�u�d<img src=card/xue1.gif>,�ڤۤ��u�b�����o�eģ�������m.����e�W�F3000�s�I,%%���I�s�I�@:)...</font>"
case "�Ѥ��d"
    conn.Execute ("update �Τ� set allvalue=allvalue+1000") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]���D�Z��׺��Ҧ����a�W�[1000�s�I�A�ڤZ��A��ܩ��߷P�¡I');}</script>"
	kapian="<img src=card/tzk.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##���N�N�o��j�a�ϥΤF�Ѥ��d<img src=card/tzk1.gif>,�Ѧa�U��,�ߧڿW�L.##���Ѥ��d�b���몺Ũ���U��[��R,�Ҧ����a�s�I�[1000.MYLOVE..�o���_��,�]�u�O##�Ҿ֦�!���j�a���n�n�P�§A�a,##�A�u�O�D�Z��ת��֯���.�j�a���I�s�I�@.</font>"
case "�����d"
    conn.Execute ("update �Τ� set ����=����+88") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]���D�Z����`���Ҧ����a�W�[88�Ӫ����I');}</script>"
	kapian="<img src=card/szk.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##���N�N�o��j�a�ϥΤF�����d<img src=card/szk1.gif>,�Ѧa�U��,�ߧڿW�L.##�������d�b���몺Ũ���U��[��R,�Ҧ����a�����[88.MYLOVE..�o���_��,�]�u�O##�Ҿ֦�!���j�a���n�n�P�§A�a,##�A�u�O�D�Z��ת��֯���.</font>"
case "�]���d"
    conn.Execute ("update �Τ� set ���A='���`',��O=��O+15000") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�Ѭӯ��d�u��.�I');}</script>"
	kapian="<img src=card/mzk.gif><bgsound src=wav/hua.wav loop=2><font color=green>�i�d���j<font color=" & saycolor & ">##���N�N�ϥΤF�p�Ӧ�D�礩���]���d<img src=card/mzk1.gif>,�Ѧa�U��,�ߧڿW�L.##���]���d�b���몺Ũ���U��[��R,�@�D����b�R���j�U��g�X���_�����~.���ǥV�v��,�Q�I�ު�,�Q������,�Ҧ��H���^�ӤF,�u�O�����D����˧ήe�o���_���F�C##�A�u�O�D�Z��׺����j�ϬP��.</font>"			
case "�|���d"
	conn.Execute ("update �Τ� set �|������=1,�|�����=date()+10  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]�ϥΤF�̰��̬��̱j�����|�������d,���ߧA�����D�Z����`�����|���աI');}</script>"
	kapian="<font color=red>�i�d���j<font color=" & saycolor & ">##�ϥΤF�ڤۤ��d�A�b�����o�eģ�������m.���ۤv�����F�@�ŷ|���A�u�O���ְ�,�֨�s�έ��i�a�I</font>"
case "�W�Ŭ��d"
 conn.Execute ("update �Τ� set �Z�\�[=�Z�\�[+200,���m�[=���m�[+200,�����[=�����[+200,���O�[=���O�[+200,��O�[=��O�[+200 where �m�W='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&zhbg_name&"]�ϥΤF�����d�A��m���\�I�Z�\�B���O�B��O�B���m�M�����W���ȦU��200�I�I�I�I');}</script>"
 kapian="<font size=2 color=green>�i�W�ťd���j<font color=" & saycolor & "><font color=red>##</font>�j�s�ۡA�ڭn��j�A�ڭn��j....�q�f�U�����X�W�Ŭ��d�A�g�@½�}�W�V�m�A�Z�\�B�����B���m�B���O�B��O�W���ȦU��<font color=red>200</font>�I�I�I�I"
case "���R�d"
	conn.Execute ("update �Τ� set �y�O=�y�O+10000  where  �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]���y�O�W���F10000�I');}</script>"
	kapian="<font color=red>�i���R�d�j<font color=" & saycolor & ">##�ڭn���R�A�ڭn���R~~�ڭn�ܦ��̬��R���H�A�n�ɹL���ŪY�A�B�w�ءA�j�a�ݳo�H�b�b���A�X��P���A���10000�y�O���F�o�I</font>"
case "�˿˥d"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�A�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �Τ� where �m�W='"&to1&"'",conn
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�H�a�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing

    end if
    if rs("�ʧO")=sex then
        kapian="<img src=card/qinqin.jpg><font color=green><bgsound src=003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥο˿˥d����:��]�O"&to1&"�P"&zhbg_name&"�O�P��!�S�R�H�F�����]�Q�ڡH</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.Execute ("update �Τ� set allvalue=allvalue+50  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]�ϥΤF�˿˥d�A�w�I�ƤW��50�I�I');}</script>"
    kapian="<img src=card/ai1.gif><font color=green><bgsound src=123.wav loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥο˿˥d<img src=card/qinqin.jpg>,�ש�p�@�H�����P"&to1&"�b��ת��j�U�ƨg��<img src=card/ai.gif>�_��!�u�s�H�r�}�ڡI�P��"&zhbg_name&"���w�I�W��50�I�I�A�ݡI��"&zhbg_name&"�ֱo�����_�ӤF�I�j�s�G<font color=red><b>���d�b��ӭ����k�O�Q��!</b></font></font>" 
case "�۲ۥd"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update �Τ� set �y�O=�y�O+500 where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]���y�O�h�F500�I');}</script>"
	kapian="<font color=red>�i�d���j<font color=" & saycolor & ">##�y�x���A�t���q�q�a�V�ۤv�l�D�w�[��%%�`�����k�h,%%�uı�߯��E���A�@�ة��ַP��W���Y�A�������ߴv�b���@�貱�}�C�P���]�H���@�ءC�y�O�W�ɤF500�I..</font>"
case "���d"
 rs.close
 rs.open "select �ʧO from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
 my_xb=rs("�ʧO")
 rs.close
 rs.open "select �ʧO from �Τ� where �m�W='"&to1&"'",conn,2,2
 to_xb=rs("�ʧO")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('���ʪB�ͧA�]��H�O���O�P���ʰڡH');}</script>"
	response.end
 end if
 conn.Execute ("update �Τ� set ��O=��O-100,���O=���O+500,�y�O=�y�O+500 where �m�W='"&to1&"'")
 conn.Execute ("update �Τ� set ��O=��O-100,���O=���O+500,�D�w=�D�w-100,allvalue=allvalue+300 where �m�W='"&zhbg_name&"'")
 kapian="<bgsound src=wav/baobao.wav loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##��%%�ϥΤF���d�A�ש�p�@�H�v���P%%�b��ת��j�U�ƨg��<img src=card/baobao.gif>�_�ӡI�u�s�H�r�}�ڡI�P��##���g��ȼW�[�F300�I,�A�ݡA��##�ֱo�T�ѤT�]�Τ���ı�I�j�s�G<font color=red><b>���d�b��ӭ����k�O�Q��!</b></font>"
case "�u�R�d"
 rs.close
 rs.open "select �ʧO from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
 my_xb=rs("�ʧO")
 rs.close
 rs.open "select �ʧO from �Τ� where �m�W='"&to1&"'",conn,2,2
 to_xb=rs("�ʧO")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('���ʪB�ͧA�R����ڡH�O���O�P���ʰڡH');}</script>"
	response.end
 end if
 conn.Execute ("update �Τ� set �D�w=�D�w+50,�y�O=�y�O+50,allvalue=allvalue+500 where �m�W='"&to1&"'")
 conn.Execute ("update �Τ� set �D�w=�D�w+50,�y�O=�y�O+50,allvalue=allvalue-500 where �m�W='"&zhbg_name&"'")
 kapian="<bgsound src=wav/zhenai.wav loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##��%%�ϥΤF�u�R�d�A�ש�p�@�H�v���P%%�b��ת��j�U�ƨg��<img src=card/zhenai.gif>���u�R��%%500�g��,�R�O�L�p���^�m�A�u�s��פH�h�����ڡI�P�����誺�D�w�M�y�O���W��50�I�A�A�ݡA��##�ֱo���_�ӤF�A�j�s�G<font color=red><b>�U���l���٭n�A�R�ڥi�H�ܡH</b></font>"
case "�Я��d"
	dim sl(4)
	sl(0)="�֯�"
	sl(1)="�]��"
	sl(2)="�I��"
	sl(3)="����"
	sl(4)="�a��"
	randomize timer
	sss=int(rnd*4)+1
	if sss=5 then sss=4
	conn.Execute ("update �Τ� set sl='"&sl(sss)&"',slsj=now()+3 where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ϥΤF["&fn1&"]�{�b�o�쯫�F["&sl(sss)&"]�ᨭ�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�����ϥΤF["&fn1&"]�A���ڡA�z�P�ڦP�b...�бo�j��["&sl(sss)&"]��ۤv�����W�K</font>"
	erase sl
case "�Ѱ��d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set �O�@=false,�ާ@�ɶ�=now() where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('���\�ѨM�F["&to1&"]���m�\���A�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�����ϥΤF�Ѱ��d���A�]�����D���@��j���n�˾`�F...</font>"
case "���`�d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set ���O=int(���O/3),��O=int(��O/3) where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]����O��O�u�ѭ�Ӫ�1/3�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ש�Ԥ��i�ԡA���X�ۤv���`�d�A�V%%���Y�W���h,%%�uı���e�@�¡A��O�B���O�l���j�b..</font>"
case "�ɦ�d"
	conn.Execute ("update �Τ� set ��O=��O+50000,���O=���O+50000  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]����O�B��O���W�[�F5�U�I�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�ɦ�d�A�o�@�U�i�O���֤F�A��O�B���O�ɺ�5�U�I�I</font>"
case "�����d"
	conn.Execute ("update �Τ� set �s��=�s��+88880000  where  �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]���s�ڤW���F8888�U�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�����d�A�ۤv���p���]���ˤ��U�F�A8888�U�r�I</font>"
case "�m�\�d"
	conn.Execute ("update �Τ� set �Z�\=�Z�\+10000  where  �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]�ϥΤF�m�\�d�A�Z�\�W��1�U�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�m�\�d�A�Z�\�i�O�j�T�פW���A�ݨӰ�פS�n���ӥ��F�I</font>"
case "�𨾥d"
	conn.Execute ("update �Τ� set ����=����+300,���m=���m+300  where  �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]�ϥΤF�𨾥d�A�����M���m�U��300�I�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�𨾥d�A�����M���m�j�T�פW���A�ݨӰ�פS�h�X�@�찪��F�I</font>"
case "�[�I�d"
	conn.Execute ("update �Τ� set allvalue=allvalue+1000  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]�ϥΤF�[�I�d�A�w�I�ƤW��1000�I�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�[�I�d�A���C�C���Ϊw�I�]�[�I�A�u�O���֮�I</font>"
case "�k�ӥd"
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
 if Instr(LCase(application("zhbg_zanli")),LCase("!"&to1&"!"))=0 then
  kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & sayscolor & ">"&zhbg_name&"��"&to1&"�ϥΤF�k�ӥd,�i��{"&to1&"}�ä��O�b�������A!�եծ��O�F�@�i�k�ӥd�I</font>"
 else
  rs.close
  rs.open "select id,�|������,����,����,����,�W���Y��,�ʧO,�n�ͦW��,�t��,�q�r from �Τ� where �m�W='" & to1 &"'",conn,2,2
  jhid=rs("id")
  hydj=rs("�|������")
  jhdj=rs("����")
  jhsf=rs("����")
  if rs("�q�r")=True then
   jhmp="�q�r��"
  else
   jhmp=rs("����")
  end if
  jhtx=rs("�W���Y��")
  sex=rs("�ʧO")
'  nowinroom=session("nowinroom")
  Application.Lock
  onlinelist=Application("zhbg_onlinelist"&nowinroom)
  onlinenum=UBound(onlinelist)
  for i=1 to onlinenum step 1
   onlinexx=split(onlinelist(i),"|")
   if onlinexx(0)=to1 then
   zhbg_zm=to1&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&jhdj&"|"&jhid&"|"&hydj&"|0"&"|"&onlinexx(9)
   onlinelist(i)=zhbg_zm
   exit for
   end if
  next
  Application("zhbg_onlinelist"&nowinroom)=onlinelist
  zhbg_zanli=split(application("zhbg_zanli"),";")
  for x=0 to UBound(zhbg_zanli)
   dxname=split(zhbg_zanli(x),"|")
   if dxname(0)="!"&to1&"!" then
    dxcz=zhbg_zanli(x)&";"
    application("zhbg_zanli")=replace(application("zhbg_zanli"),dxcz,"")
    kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & sayscolor & ">"&zhbg_name&"��"&to1&"�ϥΤF�k�ӥd,���\�a��{"&to1&"}�q�������A��^��!</font>"
    exit for
   end if
  next
  Application.UnLock
 end if
case "�Y���d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set �ɨ��ɶ�=now()  where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]�ϥΤF�Y���d�A�L����A�ɨ��F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##��b��%%���欰�ݤ��L�h�A�ϥΤF�Y���d�A%%�j�s�@�n�A�w���L�h�C�����S���F...</font>"
case "�ɨ��d"
	conn.Execute ("update �Τ� set �ɨ��ɶ�=now()  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('��["&zhbg_name&"]�ϥΤF�ɨ��d�A�{�b�ɨ����\�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##:�j�s�ۡA�ڼɡA�ڼ�....�q�f�U�����X�ɨ��d�A�ɨ����\�I</font>"
case "�o�åd"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
	conn.Execute ("update �Τ� set �Z�\=int(�Z�\/3)  where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]�ϥΤF�o�åd�A�L�Z�\�u��1/3�F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##:�j�s�ۡA�֤]���n�d�ۧڡA�ڭn�������`�I�ϥΤF�o�åd�A"&to1&"���Z�\���h�j�b....</font>"
case "���~�d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
        rs.close
        rs.open "SELECT * FROM �Τ� WHERE  �m�W='" & to1 & "'",conn,2,2
        if rs("����")="�x��" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����x���H���ϥΩ��~�d!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
        if rs("����")="�x��" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����������x���H<font color=red>[["&to1&"]]</font>�ϥΩ��~�d!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
	conn.Execute ("update �Τ� set ����='�L�ͤ�',���A='�L�ͤ�' where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('��["&to1&"]�ϥΤF���~�d�A�L�w�g�ܦ��F�L�ͤF�I');}</script>"
	kapian="<img src=card/guaishou.jpg><font color=green>�i�d���j<font color=" & saycolor & ">##:�j�s�ۡA"&to1&"�ܦ��F�L�͡A�Фj�a�p�߮@�A�p�G�Ȫ��ܽШӧڳo��A�ڥi�H�O�@�j�a....</font>"
case "�M�¥d"
	conn.Execute ("update �Τ� set ���B����=0  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('�z�ϥΤF�M�¥d�A���B���ƲM0�����I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�n�h���L�h���ɥ��r�K�K</font>"
case "�n�H�d"
	conn.Execute ("update �Τ� set �`���H=0  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('�z�ϥΤF�n�H�d�A���H�`�ƲM0�����I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF[�n�H�d]�A���H�`�Ƭ�0�F�K�K</font>"
case "�ܩʥd"
	rs.close
	rs.open "select �ʧO,�t�� FROM �Τ� WHERE �m�W='" & zhbg_name &"'",conn,2,2
	sex=rs("�ʧO")
	pl=rs("�t��")
	rs.close
    if pl="�L" then 
        if sex="�k" then
               sql="update �Τ� set �ʧO='�k' WHERE �m�W='" & zhbg_name & "'"
               xb="�}�G���k��"
         end if 
          if sex="�k" then 
              sql="update �Τ� set �ʧO='�k' WHERE �m�W='" & zhbg_name & "'"
              xb="�^�T���k��"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=zhbg_name&"�ϥ��ܩʥd��,�ש�p�@�H��,�ܦ��F"&xb&"!" 
        else
          bianxi="�ϥ��ܩʥd����!��]:"&zhbg_name&"�O���a�Ǫ��H�O,����ٷQ�ܩʧr!�o�ˤ��O�îM�F!���F�g�ٹ�"&zhbg_name&"�A�o�عD�w���a���H,�S���S��"&zhbg_name&"���ܩʥd!"
        end if
	kapian="<img src=card/bxk.jpg></td><td><font color=green>�i�d���j<font color=" & saycolor & ">##�G�o�~�Y�A�J���޳N�u���i�r�A���ܡA���ܡK�K(���O�b�ܦ��J�Y�w�A�b�@�ܩʤ�N�I�^<br>�i���G�j"&bianxi&"</font>"
case "���B�d"
	rs.close
	rs.open "select �t�� FROM �Τ� WHERE �m�W='" & zhbg_name &"'",conn,2,2
	peiou=rs("�t��")
	if peiou="�L" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�A�����r�A�ڥ��S���t���r�I');</script>"
		Response.End
	end if
	conn.Execute ("update �Τ� set �t��='�L'  where �m�W='" & zhbg_name &"'")
	rs.close
	rs.open "select �t�� FROM �Τ� WHERE �m�W='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update �Τ� set �t��='�L'  where �m�W='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">##�G�Q�e�Q��A�g�L�ۤv�@�f��Q�����A�ϥΤF���B�d,�ש�Q�n�F�P["&peiou&"]���B�F�K�K</font>"
case "�m�˥d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�A�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �Τ� where �m�W='"&to1&"'",conn,2,2
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�H�a�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("�ʧO")
    if rs("�t��")<>"�L" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥηm�˥d����:��]�O"&to1&"�O���a�Ǫ��H</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("����")="�x��" then 
           kapian="���i�H��޲z���ϥηm�˥d�K�K"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
    if rs("�t��")<>"�L" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥηm�˥d����:��]�O�ۤv�w�g�O���a�Ǫ��H�F!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("�ʧO")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥηm�˥d����:��]�O"&to1&"�P"&zhbg_name&"�O�P��!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update �Τ� set �t��='"&zhbg_name&"' where �m�W='"&to1&"'"
    conn.execute "update �Τ� set �t��='"&to1&"' where �m�W='"&zhbg_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥηm�˥d,�ש�p�@�H�����P"&to1&"�����Ұ�!</font>"
case "����d"
	rs.close
	rs.open "select ���H FROM �Τ� WHERE �m�W='" & zhbg_name &"'",conn,2,2
	peiou=rs("���H")
	if peiou="�L" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�A�����r�A�ڥ��S�����H�r�I');</script>"
		Response.End
	end if
	conn.Execute ("update �Τ� set ���H='�L'  where �m�W='" & zhbg_name &"'")
	rs.close
	rs.open "select ���H FROM �Τ� WHERE �m�W='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update �Τ� set ���H='�L'  where �m�W='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j##�G�Q�e�Q��A�g�L�ۤv�@�f��Q�����A�ϥΤF����d,�ש�Q�n�F�P["&peiou&"]����F�K�K</font>"
case "���H�d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�A�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from �Τ� where �m�W='"&to1&"'",conn,2,2
    if rs("����")="�X�a" then
		Response.Write "<script language=javascript>alert('�i"&zhbg_name&"�j�H�a�O�X�a�H�A�d���F�I�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("�ʧO")
    if rs("���H")<>"�L" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥα��H�d����:��]�O"&to1&"�O�����H���H�F</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("����")="�x��" then 
           kapian="���i�H��޲z���ϥα��H�d�K�K"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
    if rs("���H")<>"�L" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥα��H�d����:��]�O�ۤv�w�g�O�����H���H�F!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("�ʧO")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥα��H�d����:��]�O"&to1&"�P"&zhbg_name&"�O�P��!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update �Τ� set ���H='"&zhbg_name&"' where �m�W='"&to1&"'"
    conn.execute "update �Τ� set ���H='"&to1&"' where �m�W='"&zhbg_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥα��H�d,�ש�p�@�H�����P"&to1&"�������H!</font>"
case "�e���d"
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
        rs.close
        rs.open "SELECT * FROM �Τ� WHERE  �m�W='" & to1 & "'",conn,2,2
        if rs("����")="�x��" then 
            kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����x���H���ϥζe���O!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"���X��ׯS�\���e���O,��"&to1&"����F!"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update �Τ� set ���A='��',�n��=now()+3 where �m�W='" & to1 & "'"
            call boot(to1,to1&"�Q"&zhbg_name&"�ϥΤF�e���O")
        else
           kapian=kapian&mzky
        end if
case "��H�d"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
      rs.close
      rs.open "SELECT * FROM �Τ� WHERE  �m�W='" & to1 & "'",conn,2,2
      if rs("����")="�x��" then 
        kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����x���H���ϥν�H�d!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"�ϥν�H�d,���_�@�}�A���G��"&to1&"��F�X�h!"
    	  call boot(to1,to1&"�Q"&zhbg_name&"�ϥΤF��H�d")     
    	else
			kapian=mtky
    	end if
case "�V�v�d"      
  wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
  rs.close
  rs.open "SELECT * FROM �Τ� WHERE  �m�W='" & to1 & "'",conn,2,2
  if rs("����")="�x��" then
     kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����x���H���ϥΥV�v�d!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
  end if
   rs.close
   qxky=qxk(to1)
   if qxky="ok" then   
	   conn.execute "update �Τ� set �n��=now()+12/(4*144),���A='�v' where �m�W='" & to1 & "'"
	   kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥΥV�v�d!��"&to1&"�εۤF!"
	   call boot(to1,to1&"�Q"&zhbg_name&"�ϥΤF�V�v�d")
   else
   		kapian=qxky
   end if
case "�d�|�d" 
  wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
  rs.close   
  rs.open "SELECT * FROM �Τ� WHERE  �m�W='" & to1 & "'",conn,2,2
  if rs("����")="�x��" then
    kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&",�A�����x���H���ϥάd�|�d!"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("�Ȩ�")+rs("�s��")
    if yl<=10000 then
    	kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥάd�|�d����:��]"&to1&"���W���Ȥl�p��10000��!"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&yl&" where �m�W='"&zhbg_name&"'"
      if rs("�Ȩ�")>=rs("�s��") then
        conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&yl&" where �m�W='"&to1&"'"
      else
        conn.execute "update �Τ� set �s��=�s��-"&yl&" where �m�W='"&to1&"'"
      end if
      kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"�ϥάd�|�d,�d�o"&to1&"�@"&yl&"��Ȥl,�����k"&zhbg_name&"�Ҧ�!"
    else
       kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"��"&to1&"�ϥάd�|�d����!"
       kapian=kapian&mhky
    end if
case "�ɯťd"
 rs.close
 rs.open "select allvalue,�|������,���� from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
 if rs("�|������")>3 or rs("����")>100 then
		Response.Write "<script language=javascript>alert('���ܡG100�ťH�W��4�ťH�W�|���L�ݨϥΤɯťd�I');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 jhjy=rs("allvalue")
 jhdj=int(sqr(jhjy/50))
 jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
 jhdj=jhdj+1
 jhjy=jhjy+jhadd
 conn.Execute ("update �Τ� set allvalue="&jhjy&",����="&jhdj&" where �m�W='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&zhbg_name&"]�ϥΤF�ɯťd�I��׸g��ȤW�ɤF"&jhadd&"�I�A�԰����ŤW��1�šA�{��"&jhdj&"��');}</script>"
 kapian=zhbg_name&"�ϥΤF�ɯťd�A"&zhbg_name&"���w�I�W�["&jhadd&"�I�A�԰����ŤW��1��...�u�O�n�֮�r�A���w�I�]�ɯšI"
case "�����d"
 conn.Execute ("update �Τ� set �Z�\�[=�Z�\�[+500,���O�[=���O�[+500,��O�[=��O�[+500 where �m�W='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('��["&zhbg_name&"]�ϥΤF�����d�A��m���\�I�Z�\�B���O�B��O�W���ȦU��500�I�I�I�I');}</script>"
 kapian="�i�d���j"&zhbg_name&"�j�s�ۡA�ڭn��j�A�ڭn��j....�q�f�U�����X�����d�A�b"&Application("zhbg_user")&"����߽ձФU�A"&zhbg_name&"�g�L�@½�}�W�V�m�A�Z�\�B���O�B��O�W���ȦU��500�I�I�I�I"
case "�n�H�d"
	conn.Execute ("update �Τ� set ���H��=0  where �m�W='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('�z�ϥΤF�n�H�d�A���H�ƲM0�����I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##���A�ڬO�n�H�A�ڤ��ѨS�����H�r�K�K</font>"
case "�Ѳ��d"
	Response.Write "<script>parent.slbox=true;<"&"/script>"	
	Response.Write "<script language=JavaScript>{alert('�z�ϥΤF�Ѳ��d�A�{�b�i�H�ݨ�O�H�p��F�I');}</script>"
	kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�����a�ѱҤF�Ѳ��A�Q�U�ʤd�����a�賣��ݨ��F�K�K</font>"
case "���ťd"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
     rs.close
     rs.open "select * from �Τ� where �m�W='"&to1&"'",conn,2,2
     jhdj=rs("����")
     del1=jhdj*jhdj*50
     del2=(jhdj+1)*(jhdj+1)*50
     delok=del2-del1
     jhjf=rs("allvalue")-delok
 	conn.Execute ("update �Τ� set ����=����-1,allvalue="&jhjf&" where �m�W='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]�����ŭ���"&jhdj-1&"�ŤF�I');}</script>"
	kapian="<img src=card/jiangji.jpg><font color=green><bgsound src=wav/003.mid loop=1><font color=green>�i�d���j<font color=" & saycolor & ">##��%%�ϥΤF���ťd,%%�����ŭ���"&jhdj-1&"�šA�n���]�j�����.</font>"
case "���ץd"
rs.close
rs.open "select * from �Τ� where �m�W='"&to1&"'",conn,2,2
if rs("�q�r")="True" then
kapian="<bgsound src=plus/wav/wav/003.mid loop=1><font color=green>�i�d���j</font><font color=blue>"&zhbg_name&"</font>��<font color=green>"&to1&"</font>�ϥ�<font color=red><b>���ץd</b></font>���ѡG"&to1&"�]�O�q�r�ǡA�w�g���ΧA���פF�I�I"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
if rs("����")="�x��" then 
  kapian=""&zhbg_name&"�A���i�H��޲z���ϥζ��ץd�K�K"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
rs.close
rs.open "select �q�r from �Τ� where �m�W='"&zhbg_name&"'",conn,2,2
if rs("�q�r")="False" then
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>�i�d���j</font><font color=blue>"&zhbg_name&"</font>��<font color=green>"&to1&"</font>�ϥ�<font color=red><b>���ץd</b></font>���ѡG"&zhbg_name&"���O�q�r�ǡA�����קr�I�I"
        rs.close
set rs=nothing
conn.close
set conn=nothing
exit function
end if
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##������%%�ϥΤF["&fn1&"]...</font>"&wnky
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
exit function
end if
conn.execute "update �Τ� set �q�r=True where �m�W='"&to1&"'"
conn.execute "update �Τ� set �q�r=False where �m�W='"&zhbg_name&"'"
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>�i�d���j</font>##�����ϥΥd���G�{�b%%�O�q�r�Ǥj�a�ֱ��r�I"
Case "����d"
    rs.close
    rs.Open "select pig from �Τ� where  �m�W='" & zhbg_name & "'", conn,2,2
    myhua = rs("pig")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>�i�d���j</font>##�A�èS���p��,����ϥν���d�A�Хh�س]�ް�!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>�i�d���j</font>##�A������ƾڦ����D�A�Э��s�س]�ް�!"
                Exit Function
        End If
        '�W�[����
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update �Τ� set pig='" & newmyhua & "' where �m�W='" & zhbg_name & "'"
        kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF����d<img src=card/zhuzai.JPG><bgsound src=mav/XL.MID loop=1>,�o��F5�Y����A�֥h�i�ާa�I</font>"
Case "�ؤl�d"
    rs.close
    rs.Open "select hua from �Τ� where  �m�W='" & zhbg_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>�i�d���j</font>##�A�èS���A��,����ϥκؤl�d�A�Хh�}��!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>�i�d���j</font>##�A���A��ƾڦ����D�A�Э��s�}��!"
                Exit Function
        End If
        '�W�[�ؤl
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update �Τ� set hua='" & newmyhua & "' where �m�W='" & zhbg_name & "'"
        kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�ؤl�d,�o��F5����ءA�֥h�ت�a�I</font>"
Case "�ت�d"
    rs.close
    rs.Open "select hua from �Τ� where  �m�W='" & zhbg_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>�i�d���j</font>##�A�èS���A��,����ϥκت�d�A�Хh�}��!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>�i�d���j</font>##�A���A��ƾڦ����D�A�Э��s�}��!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [�Τ�] set hua='" & newmyhua & "' where �m�W='" & zhbg_name & "'"
        kapian="<font color=green>�i�d���j<font color=" & saycolor & ">##�ϥΤF�ت�d,�u�O�n�ڡA���ݵ۪ᨳ�t�a�W���I</font>"
case else
	Response.Write "<script language=JavaScript>{alert('�t�ΨèS��["&fn1&"]�o�إd��,�Τ���ϥΧA�d���F�I');}</script>"
	Response.End
end select
'�R���ۤv�d���A�O��
conn.execute "update �Τ� set w5='"&mycard&"' where  �m�W='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','�ާ@','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'�K�|�d
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from �Τ� where �m�W='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"�K�|�d")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"�K�|�d",1)
		conn.execute "update �Τ� set w5='"&tocard&"' where  �m�W='"&to1&"'"
	   'mhk="<br><font color=green>�i�K�|�d�j</font>"&to1&"���W���K�|�d�ͮ�,�]�������L!"
	   mhk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"���ǳƳߴ������q"&to1&"���f�U������,�N�b����,"&to1&"�ǥX���W���K�|�d��,�C��,�ڨ��W���K�|�d,�K�K!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'�K�o�d
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from �Τ� where �m�W='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"�K�o�d")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"�K�o�d",1)
		conn.execute "update �Τ� set w5='"&tocard&"' where  �m�W='"&to1&"'"
	   'mzk="<br><font color=green>�i�K�o�d�j</font>"&to1&"���W���K�o�d�ͮ�,�]�������L!"
	   mzk="<img src=card/myk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&"�x�����H�ǳƥή��K���M�b"&to1&"����l�W,�N�b����,"&to1&"�ǥX���W���K�o�d��,�C��,�ڨ��W���K�o�d,�K�K!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'�M���d
function qxk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from �Τ� where �m�W='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"�M���d")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"�M���d",1)
		conn.execute "update �Τ� set w5='"&tocard&"' where  �m�W='"&to1&"'"
	   qxk="><img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"���X�����y�A�Χa�A�Χa�A�b�ʯv�K�K"&to1&"�C�ۭӤj�����A�̼H�H���ݵۥL�K�K�b���ڶܡH"&to1&"�ǥX���W���M���d,�C��,�ڨ��W���M���d,�K�K!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'�K��d
function mtk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from �Τ� where �m�W='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"�K��d")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"�K��d",1)
		conn.execute "update �Τ� set w5='"&tocard&"' where  �m�W='"&to1&"'"
	   mtk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&zhbg_name&"�ϥX�겣��}�A�ǳƨӭӰ�ڦ�ʡA�o���p�߽����Y�K�K"&to1&"�b�@��K�K�����A�N�A�r�A�٭n��ڡA�A��20�~�K�K"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'�U��d
function wnk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from �Τ� where �m�W='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"�U��d")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"�U��d",1)
		conn.execute "update �Τ� set w5='"&tocard&"' where  �m�W='"&to1&"'"
	   wnk="<img src=card/wangneng.jpg></td><td><font color=green style='font-size=9pt'><bgsound src=wav/003.mid loop=1>�i�d���j<font color=" & saycolor & ">"&to1&"�b�@��K�K�����A�N�A�r�A�U��d�A�U��d�A�@�d�b��A���M�ѤU�K�K"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>