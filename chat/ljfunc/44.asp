<%'�d��
function kapian(fn1,to1,toname)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
if toname="�j�a"  then
	Response.Write "<script language=JavaScript>{alert('�Υd�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(0)=toname and instr(";�Ѱ��d;���`�d;�d�|�d;�t�¥d;���F�d;�m�˥d;�𨧥d;�o�åd;�ݮ`�d;��H�d;�e���d;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('�i"&fn1&"�j����ۤv�i��ާ@�I');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
fn1=trim(fn1)
rs.open "select ���~�W FROM ���~ WHERE ����='�d��'  and �ƶq>0 and ���~�W='" & fn1 & "' and �֦���="& info(9),conn
if rs.eof or rs.bof  then
	Response.Write "<script language=JavaScript>{alert('�A�èS��["&fn1&"]�o�إd��,�Y�Q�A���ϥ�,�бz�ʶR�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

select case fn1
case "�t�¥d"
conn.Execute ("update �Τ� set kill=0 where �m�W='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('���\��["&toname&"]�ϥΤF�t�¥d�I');}</script>"
kapian=info(0)&"��["&toname&"]�ϥΤF�t�¥d���A�Ϲ�誺�Q�����Ƭ�0,�~����`�F..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�Ѱ��d"
conn.Execute ("update �Τ� set �O�@=false,�ާ@�ɶ�=now() where �m�W='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('���\�ѨM�F["&toname&"]���m�\���A�I');}</script>"
kapian=info(0)&"�����ϥΤF�Ѱ��d���A�]�����D���@��j���n�˾`�F..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "���F�d"
rs.close
rs.open "select �s�� FROM �Τ� WHERE �m�W='"&toname&"'",conn
yin=rs("�s��")
yin=yin/2
conn.Execute ("update �Τ� set �s��=�s��+" & yin & "  where id="&info(9))
conn.Execute ("update �Τ� set �s��=�s��-" & yin & "  where �m�W='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('�A���\�����設�W���s�ڷm���F1/2�I');}</script>"
rs.close
rs.open "select �s�� FROM �Τ� WHERE �m�W='" & toname & "'",conn
if rs("�s��")<0 then 
conn.Execute ("update �Τ� set �s��=0 where �m�W='" & toname & "'")
end if
kapian=info(0)&"��["&toname&"]�ϥΤF���F�d�A���\�����設�W���s���F��1/2,�@�p"&yin&"..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�m�˥d"
rs.close
rs.open "select �t�� FROM �Τ� WHERE id="&info(9),conn
yin=rs("�t��")
if yin<>"�L"  then
	Response.Write "<script language=JavaScript>{alert('�A���ѱC�F�r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select �t�� FROM �Τ� WHERE �m�W='"&toname&"'",conn
yina=rs("�t��")
if yina="�L" then
	Response.Write "<script language=JavaScript>{alert('����٨S�ѱC�r�A�A�n�m����r�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	conn.Execute ("update �Τ� set �t��='" & yina & "'  where id="&info(9))
    conn.Execute ("update �Τ� set �t��='�L'  where �m�W='" & toname & "'")
conn.Execute ("update �Τ� set �t��='" & info(0) & "'  where �m�W='" & yina & "'")
	Response.Write "<script language=JavaScript>{alert('�A���\�����誺�ѱC�m�F�L�ӡA�����I');}</script>"
kapian=info(0)&"��["&toname&"]�ϥΤF�m�˥d�A���\����["&toname&"]���ѱC["&yina&"]���m�F�L��..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�{�q�d"
	conn.Execute ("update �Τ� set ���O=���O/2  where �m�W='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('�A���\���ϥΤF�{�q�d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�{�q�d�A�Ϲ�誺���O��b..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�d�|�d"
	conn.Execute ("update �Τ� set �s��=�s��*0.7 where �m�W='" & toname & "'")
    conn.Execute ("update �Τ� set �Ȩ�=�Ȩ�*0.7 where �m�W='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('���\������ϥάd�|�d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�d�|�d�A�Ϲ�設�W���Ȩ�M�Ȧ�s�ڳQ����30%�|��..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")

case "�m�ܥd"
rs.close
rs.open "select �m�W,�Ȩ� FROM �Τ� WHERE �m�W='"&toname&"'",conn
yin=int(rs("�Ȩ�")*0.5)
conn.Execute ("update �Τ� set �Ȩ�=�Ȩ�+" & yin & "  where id="&info(9))
conn.Execute ("update �Τ� set �Ȩ�=�Ȩ�-" & yin & "  where �m�W='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('�A���\�����設�W���Ȥl�m���F1/2�I');}</script>"
rs.close
rs.open "select �Ȩ� FROM �Τ� WHERE �m�W='" & toname & "'",conn
if rs("�Ȩ�")<0 then 
conn.Execute ("update �Τ� set �Ȩ�=0  where �m�W='" & toname & "'")
end if
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
kapian=info(0)&"��"&toname&"�ϥΤF�m�ܥd�A���\�����設�W���Ȥl�m���F3/4,�@�p"&yin&"..."
case "���`�d"
conn.Execute ("update �Τ� set ���O=int(���O/3),��O=int(��O/3) where �m�W='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('["&toname&"]����O��O���ѭ�Ӫ�1/3�F�I');}</script>"
kapian=info(0)&"�ש�Ԥ��i�ԡA���X�ۤv���`�d�A�V"&toname&"���Y�W���h,"&toname&"��ı���e�@�¡A��O�B���O�l���j�b.."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�k�N�d"
conn.Execute ("update �Τ� set �k�O=�k�O+2000 where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���k�O�W�[2000�I�I');}</script>"
kapian=info(0)&"�ϥΤF�k�N�d�A�o�U�i�n�F�A�k�O�@�ɤ����W�[�F2000�I�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�ǭ��d"
conn.Execute ("update �Τ� set �O�@=false,���H��=0,�ާ@�ɶ�=now() where �|������=1")
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\�ϥΩǭ��d�I');}</script>"
kapian="<center><font size=4 color=red>�i"&info(0)&"�����ϥΤF�ǭ��d�A1�ŷ|�����O�@���Ѱ��F...�j</center><bgsound src=wav/003.MID loop=1></font>"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")

case "�ɦ�d"
	conn.Execute ("update �Τ� set ��O=��O+50000,���O=���O+50000  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]����O�B��O���W�[�F5�U�I�I');}</script>"
kapian=info(0)&"�ϥΤF�ɦ�d�A�o�@�U�i�O���֤F�A��O�B���O�ɺ�5�U�I�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�}���d"
			conn.Execute ("update �Τ� set �Ȩ�=0  where �m�W='" & toname & "'")
			Response.Write "<script language=JavaScript>{alert('�A���\���ϥΤF�}���d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�}���d�A�Ϩ䨭�W���Ȩ�@�����S���F�A�n�G�@..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�����d"
conn.Execute ("update �Τ� set �s��=�s��+88880000  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���s�ڤW���F8888�U�I');}</script>"
kapian=info(0)&"�ϥΤF�����d<img src=card/mhk.jpg>�A�ۤv���p���]���ˤ��U�F�A8888�U�r�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�m�\�d"
conn.Execute ("update �Τ� set �Z�\=�Z�\+10000  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�ϥΤF�m�\�d�A�Z�\�W��1�U�I');}</script>"
kapian=info(0)&"�ϥΤF�m�\�d�A�Z�\�i�O�j�T�פW���A�ݨӦ���S�n���ӥ��F�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�Ѫ�d"
	conn.Execute ("update �Τ� set ��O=��O/2  where �m�W='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('�A���\���ϥΤF�Ѫ�d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�Ѫ�d�A�Ϲ�誺��O��b..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�_��d"
conn.Execute ("update �Τ� set ���H��=0  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�����H�ƫ�_���l���A�A�S�i�H�}�l���H�F�I');}</script>"
kapian=info(0)&"�ϥΤF�_��d�A�Ϧۤv�����H�ƫ�_���l���A�A�S�i�H�}�l���H�F�r�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�[�I�d"
	conn.Execute ("update �Τ� set allvalue=allvalue+1000  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�ϥΤF�[�I�d�A�w�I�ƤW��1000�I�I');}</script>"
kapian=info(0)&"�ϥΤF�[�I�d�A���C�C���Ϊw�I�]�[�I�A�u�O���֮�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�]���d"
yin2=100000*Application("ljjh_chatrs"&session("nowinroom"))
conn.Execute ("update �Τ� set �Ȩ�=�Ȩ�+" & yin2 & " where �m�W='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('�A���\���ϥΤF�]���d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�]���d�A�Ϩ䦬���ѫǦb�u�H�ƭ�10�U�⪺�Ȥl..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�I���d"
rs.close
rs.open "select �Ȩ� FROM �Τ� WHERE �m�W='" & toname & "'",conn
yin3=rs("�Ȩ�")/Application("ljjh_chatrs"&session("nowinroom"))
	useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
			online=Split(Trim(useronlinename)," ",-1)
			x=UBound(online)
			for i=0 to x
conn.Execute ("update �Τ� set �Ȩ�=�Ȩ�+" & yin3 & " where �m�W='" & online(i) & "'")
			next
conn.Execute ("update �Τ� set �Ȩ�=0 where �m�W='" & toname & "'")
			Response.Write "<script language=JavaScript>{alert('�A���\���ϥΤF�I���d�I');}</script>"
kapian=info(0)&"��"&toname&"�ϥΤF�I���d�A��ѫǦb�y���C�ӤH���o"&toname&"��"&yin3&"��Ȥl�A�ݥL�I�����ˡA��������~~~�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�𨧥d"
	conn.Execute ("update �Τ� set �ɨ��ɶ�=now()  where �m�W='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('��["&toname&"]�ϥΤF�𨧥d�A�L����A�ɨ��F�I');}</script>"
kapian=info(0)&"��b��"&toname&"���欰�ݤ��L�h�A�ϥΤF�𨧥d�A"&toname&"�j�s�@�n�A�w���L�h�C�����S���F..."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")

case "�ɨ��d"
	conn.Execute ("update �Τ� set �ɨ��ɶ�=now()  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('��["&info(0)&"]�ϥΤF�ɨ��d�A�{�b�ɨ����\�F�I');}</script>"
kapian=info(0)&"�j�s�ۡA�ڼɡA�ڼ�....�q�f�U�����X�ɨ��d�A�ɨ����\�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�D�w�d"
	conn.Execute ("update �Τ� set �D�w=�D�w+100000  where id="&info(9))
rs.close
rs.open "select ����,�D�w,�D�w�[ FROM �Τ� WHERE id="&info(9),conn
ddj=(rs("����")*1440+2200)+Rs("�D�w�[")
if rs("�D�w")>ddj then
conn.Execute ("update �Τ� set �D�w=" & ddj & "  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���D�w�W���F10�U�I');}</script>"
        kapian=info(0)&"�ϥΤF�D�w�d�A�ۤv���D�w���_�W���A���F10�U�@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�o�åd"
	conn.Execute ("update �Τ� set �Z�\=int(�Z�\/3)  where �m�W='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('��["&toname&"]�ϥΤF�o�åd�A�L�Z�\����1/3�F�I');}</script>"
kapian=info(0)&"�j�s�ۡA�֤]���n�d�ۧڡA�ڭn�������`�I�ϥΤF�o�åd�A"&toname&"���Z�\���h�j�b...."
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�y�O�d"
	conn.Execute ("update �Τ� set �y�O=�y�O+200000  where id="&info(9))
rs.close
rs.open "select ����,�y�O,�y�O�[ FROM �Τ� WHERE id="&info(9),conn
mlj=(rs("����")*1120+2100)+rs("�y�O�[")
if rs("�y�O")>mlj then
conn.Execute ("update �Τ� set �y�O=" & mlj & "  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���y�O�W���F20�U�I');}</script>"
        kapian=info(0)&"�ϥΤF�y�O�d�A�ۤv���y�O���_�W���A���F20�U�@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")

case "�����d"
rs.close
rs.open "select �D�w FROM �Τ� WHERE id="&info(9),conn
if rs("�D�w")<0 then
conn.Execute ("update �Τ� set �D�w=10000  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���D�w��_���`�I');}</script>"
        kapian=info(0)&"�ϥΤF�����d�A�ۤv���D�w��_�F���`�A�i�H�����c�H�F�@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "���B�d"
randomize timer
r=int(rnd*1200)+800
rs.close
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
conn.Execute ("update �Τ� set ����=����+"&r&"  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]������u�n�A�����W�[�F"&r&"�ӡI');}</script>"
        kapian=info(0)&"�ϥΤF���B�d�A������M�W�[�F"&r&"�ӡA�����٤��t�@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�ݮ`�d"
rs.close
rs.open "select ��O FROM �Τ� WHERE  �m�W='" & toname & "'",conn
conn.Execute ("update �Τ� set ��O=1000  where �m�W='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF�ݮ`�d�I');}</script>"
        kapian=info(0)&"��" & toname & "�ϥΤF�ݮ`�d�A" & toname & "�Q�ѥ~���L�@�y���~�A��O�ȳ�1000�F�@,�Ѧ�j�L���ɦA���U���ݦ�ɮ@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�D§�d"
rs.close
rs.open "select ��O FROM �Τ� WHERE  id="&info(9),conn
conn.Execute ("update �Τ� set �y�O=�y�O-1000  where �m�W='" & toname & "'")
conn.Execute ("update �Τ� set �y�O=�y�O+10000,����=����+10000  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF�D§�d�I');}</script>"
        kapian=info(0)&"��" & toname & "�ϥΤF�D§�d�A" & toname & "�Q"&info(0)&"�������A�@�ɮ�w�L�h�A�y�O���1000�A"&info(0)&"�@�A�A�ӥi�c�����F�j,"&info(0)&"�D�w���1000�I�A�y�O�B���W�[10000�I�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�[���d"
rs.close
rs.open "select �w���I�� FROM �Τ� WHERE  id="&info(9),conn
conn.Execute ("update �Τ� set �w���I��=�w���I��+10000  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF�[���d�I');}</script>"
        kapian=info(0)&"��ۤv�ϥΤF�[���d�A�߮ɤf�U�̨����h�h�A�ּ֦h�h�@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case "�s�`�d"
	Response.Write "<script language=JavaScript>{alert('�s�`�d�L�ݨϥΡA���ݫO�d�I');}</script>"
Response.End
case "�j�K�d"
	Response.Write "<script language=JavaScript>{alert('�j�K�d�L�ݨϥΡA���ݭn�O�d�I');}</script>"
Response.End
case "��H�d"
rs.close
rs.open "select grade FROM �Τ� WHERE  �m�W='" & toname & "'",conn
if rs("grade")>5 then
	Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ާ@�θӤH�O�S�O�O�@�ﹳ!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF��H�d�I');}</script>"
        kapian=info(0)&"��" & toname & "�ϥΤF��H�d�A����" & toname & "�@�ӵ����˭��X�F��ѫǮ@�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
call boot(toname)
case "�e���d"
rs.close
rs.open "select grade FROM �Τ� WHERE  �m�W='" & toname & "'",conn
if rs("grade")>5 then
	Response.Write "<script language=JavaScript>{alert('���ѡA�A����ﰪ�ź޲z���ާ@�θӤH�O�S�O�O�@�ﹳ!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select id FROM �Τ� WHERE �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select ���~�W FROM ���~ WHERE ����='�d��'  and �ƶq>0 and ���~�W='�j�K�d' and �֦���="& idd,conn
if not (rs.eof or rs.bof)  then

conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&idd&" and ���~�W='�j�K�d'")
	Response.Write "<script language=JavaScript>{alert('��設�W�{���j�K�d�A�е��L�����|���n��L�j�A�������ӭ��lOK�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.Execute ("update �Τ� set ���A='��',�n��=now()+3  where �m�W='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF�e���d�I');}</script>"
        kapian=info(0)&"��" & toname & "�ϥΤF�e���d�A����" & toname & "�Y�W�_�X�@���¤�@��N��" & toname & "���F�Ѩc�I"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
call boot(toname)
case "�ܩʥd"
rs.close
rs.open "select �ʧO,�t��,�G�B FROM �Τ� WHERE  id="&info(9),conn
xb=rs("�ʧO")
pei=rs("�t��")
er=rs("�G�B")
if xb="�k" then
xb="�k"
else
xb="�k"
end if
conn.Execute ("update �Τ� set �ʧO='"&xb&"',�t��='�L',�G�B='�L'  where id="&info(9))
conn.Execute ("update �Τ� set �t��='�L'  where �m�W='"&pei&"'")
conn.Execute ("update �Τ� set �G�B='�L'  where �m�W='"&er&"'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]���\���ϥΤF�ܩʥd�I');}</script>"
        kapian=info(0)&"��ۤv�ϥΤF�ܩʥd�A����"&info(0)&"�j�s�G���ܡA���ܡA��������,"&info(0)&"���\����ۤv��y���@��???"
conn.Execute ("update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9) &" and ���~�W='"&fn1&"'")
case else
	Response.Write "<script language=JavaScript>{alert('�A�èS��["&fn1&"]�o�إd��,�Y�Q�A���ϥ�,�бz�ʶR�I');}</script>"
	Response.End
end select
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','�ާ@','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>