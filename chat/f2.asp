<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
id=request.querystring("id")
%>
<html><head><meta http-equiv='content-type' content='text/html; charset=big5'>
<style type='text/css'>
.webstyle   {font-family: Webdings; font-size: 9pt}
.yy4{ font-size: 9pt;color:FFFFFF; line-height: 11pt;position: relative; width: 100% }
body{font-size:9pt;}input{font-size:9pt;}a{font-size:9pt;color:FFFF00;text-decoration:none;}a:hover{color:FFFF00;text-decoration:underline;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=000000 background=http://xs68.xs.to/pics/06072/b39.gif  bgproperties=fixed topmargin=0 text=FFFF00>
<form name=af method=POST action='say.asp'  target='d' onsubmit='return(parent.checksays());'>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td>
<div align=center><input type=hidden name='mdtx' value="<%=info(10)%>">
<input type=hidden name='grade' value="<%=info(2)%>">
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<select name='addwordcolor' style='font-size:9pt;background-color:#000000'>
<option style="color:00FF00" value="00FF00" selected>�W</option>
<option style="color:00FF00" value="00FF00">�W</option>
<option style="color:0000FF" value="0000FF">�W</option>
<option style="color:FFFF00" value="FFFF00">�W</option>
<option style="color:FFCC00" value="FFCC00">�W</option>
<option style="color:FFD0CF" value="FFD0CF">�W</option>
<option style="color:F0E0A0" value="F0E0A0">�W</option>
<option style="color:B0DF90" value="B0DF90">�W</option>
<option style="color:90CFEF" value="90CFEF">�W</option>
<option style="color:CF9FC0" value="CF9FC0">�W</option>
<option style="color:C0C0C0" value="C0C0C0">�W</option>
<option style="color:FFFFFF" value="FFFFFF">�W</option>
<option style="color:A992D6" value="A992D6">�W</option>
<option style="color:53A6A6" value="53A6A6">�W</option>
</select>
<select name='sayscolor'  style='font-size:9pt;background-color:#000000'>
<option style="color:FFCC00" value="FFCC00" selected>��</option>
<option style="color:00FF00" value="00FF00">��</option>
<option style="color:0000FF" value="0000FF">��</option>
<option style="color:FFFF00" value="FFFF00">��</option>
<option style="color:FFCC00" value="FFCC00">��</option>
<option style="color:F0E0A0" value="F0E0A0">��</option>
<option style="color:B0DF90" value="B0DF90">��</option>
<option style="color:90CFEF" value="90CFEF">��</option>
<option style="color:CF9FC0" value="CF9FC0">��</option>
<option style="color:C0C0C0" value="C0C0C0">��</option>
<option style="color:FFFFFF" value="FFFFFF">��</option>
<option style="color:A992D6" value="A992D6">��</option>
<option style="color:53A6A6" value="53A6A6">��</option>
</select>
<select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>��</option>
<option value="�L���a">�L��</option>
<option value="�ŬX�a">�ŬX</option>
<option value="�����y">�y��</option>
<option value="�n�Y�̸��o�N�a">�o�N</option>
<option value="���I���I���I����">�j��</option>
<option value="�֭n���a">�֭�</option>
<option value="����">��</option>
<option value="���h�n�N�a">�a�N</option>
<option value="��Ѧa">���</option>
<option value="���j�F�����A�ܸҲ��a">�Ҳ�</option>
<option value="���֦a">����</option>
<option value="�d�h�a">�d�h</option>
<option value="���q���M�a">���q</option>
<option value="�Y�¦a">�Y��</option>
<option value="�ͮ�a">�ͮ�</option>
<option value="�j�n�a">�j�n</option>
<option value="�̥G�G�a">��</option>
<option value="�ܺ����a">����</option>
<option value="�ܵL�d�a">�L�d</option>
<option value="�̨̤��٦a">����</option>
<option value="�f�R�ժj">�ժj</option>
<option value="�L�i�`��a">�L�`</option>
<option value="�i�������a">�i��</option>
</select>
<input type=text name='username' readonly  style="border:1px solid #6699FF; text-align:center;font-size:9pt;color:#FFCCCC;background-color:#000000" size=5 maxlength=10>��
<select name='towho' style='font-size:12px;color:#FFFFFF;background-color:000000' onClick=dj()><option value='�j�a' selected>�j�a</option></select>
<input type=text name='sytemp'   style='font-size:9pt;color:ffffff;background-color:#000000' size=30 maxlength=180 >
<a href="#" onClick="gop()"><img src=ico/wfy_i_c_g_b32.gif border=0></a>
<a href="#" onClick="gon()"><img src=ico/wfy_i_c_g_b31.gif border=0></a>
<input type=submit  name='subsay' value='�o��' style="background-color:FF0000;color:#FFFFFF;border: 1 double" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">
<input type=text name='clock' style="text-align:right;font-size:9pt;background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " value="" size=4 readonly>
</td></tr>
</table>
<div align=center>
<input name="tu" type="hidden" value="">
<input name="addsign" type="hidden" value=0>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option selected>�߱�
<option style=background-color:"#000000" value="/�߱�$ ���`">���`
<option style=background-color:"#000000" value="/�߱�$ �w�I">�w�I
<option style=background-color:"#000000" value="/�߱�$ �۩w">�۩w
</select>
<select name='command' onchange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>�ϥΩR�O1 </option>
<option style="color:#FFFFFF" value="/�H��$">�d�ݫH�� </option>
<option style="color:#FFFFFF" value="/�m��$">�m���_�� </option>
<option style="color:#FFFFFF" value="/��D$ �D�Ԥ��">��Z��D </option>
<option style="color:#FFFFFF" value="/�D�B$ �u�R���">�ܷR�D�B </option>
<%if info(1)="�k" then %>
<option style="color:#FFFFFF" value="/�D�c$">�D�p�ѱC </option>
<% else %>
<option style="color:#FFFFFF" value="/�D�p�Ѥ�$">�D�p�Ѥ� </option>
<%end if%>
<option style="color:#FFFFFF" value="/�P�N$">�����L�� </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�D�D$ ">�d�̶ǭ� </option>
<option style="color:#FFFFFF" value="/�߰�$ ">�߰ʷPı </option>
<option style="color:#FFFFFF" value="/��q$ ">�g���q </option>
<option style="color:#FFFFFF" value="/�߸�$ ">�߸��߻y </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/���$">������� </option>
<option style="color:#FFFFFF" value="/�׽m$ ">�_���׽m </option>
<option style="color:#FFFFFF" value="/�ɨ�$ ">�ɫ¤O�� </option>
<option style="color:#FFFFFF" value="/���v$ ">���v�ߪZ </option>
<option style="color:#FFFFFF" value="/�U�s$ ">�Ǧ��U�s </option>
<option style="color:#FFFFFF" value="/���{$ ">�ۦ��{�� </option>
<option style="color:#FFFFFF" value="/�[�J$ �����W">�[�J����</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�Ǳ�$ 100">�Ǳ¤��O </option>
<option style="color:#FFFFFF" value="/�g��$ 100">�Ǳ¸g�� </option>
<option style="color:#FFFFFF" value="/�e��$ 1000">�ذe�{�� </option>
<option style="color:#FFFFFF" value="/�ذe$ �g�X���~�W">�ذe���~ </option>
<option style="color:#FFFFFF" value="/���$ �g�X���~�W">��󪫫~ </option>
<option style="color:#FFFFFF" value="/�ϥ�$ �d���W">�ϥΥd�� </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/����$ ">����^�K</option>
<option style="color:#FFFFFF" value="/���e$ ">���e�j�k</option>
<option style="color:#FFFFFF" value="/�ɮ�$ ">�C���F��</option>
<option style="color:#FFFFFF" value="/�ЪZ$ ">�оɪZ�\</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/����">�������</option>
<option style="color:#FFFFFF" value="/��F�o�a$ ">�����</option>
<option style="color:#FFFFFF" value="/�p��$ ">�������~</option>
<option style="color:#FFFFFF" value="/��p��$ ">��p��</option>
<option style="color:#FFFFFF" value="/�Ĥl����$ ">�Ĥl����</option>
<option style="color:#FFFFFF" value="/�аV�Ĥl$ ">�аV�Ĥl</option>
</select>
<select name='command2' onChange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>�ϥΩR�O2 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�����J$ 1000|�Ȩ�[�ΡG�����A�k�O]">�����J
<option style="color:#FFFFFF" value="/�o�P$ ����|�t[�γ�]" style=color:#000000>������
<option style="color:#FFFFFF" value="/���±N$ 1000|�Ȩ�[�ΡG�����A�k�O]">���±N
<option style="color:#FFFFFF" value="/�X�P$ ����|�t[�γ�]" style=color:#000000>�¤���
<option style="color:#FFFFFF" value="/���B$ ���B�Ʀr1-1000">�F�C�m�� </option>
<option style="color:#FFFFFF" value="/���H���$ �ж�J�Ʀr">��H��� </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�o��$ �ۦ��W">�o�ۧ���
<option style="color:#FFFFFF" value="/�l�P�j�k$ ">�l�����O </option>
<option style="color:#FFFFFF" value="/����$ ">�������] </option>
<option style="color:#FFFFFF" value="/�U�r$ �r�ĦW">�����U�r </option>
<option style="color:#FFFFFF" value="/���Y$ �t���W">���Y�t�� </option>
<option style="color:#FFFFFF" value="/���R$ ">�P�k��� </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/����$ ">�V�W���� </option>
<option style="color:#FFFFFF" value="/���I$ ">�s�}���I </option>
<option style="color:#FFFFFF" value="/����$ ">�����m�\ </option>
<option style="color:#FFFFFF" value="/����$ ">���ؾi�� </option>
<option style="color:#FFFFFF" value="/�m�Z$ ">�Z�]�ߪZ </option>
<option style="color:#FFFFFF" value="/�ײߪZ�\$ �Z�\�ۦ�">�ײߪZ�\ </option>
<option style="color:#FFFFFF" value="/�׾i$ ">�פ߾i�� </option>
</select>
<select name='command3' onChange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>�ϥΩR�O3</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�t���_��$ ">�t���_��</option>
<option style="color:#FFFFFF" value="/�]�O�p��$ ">�]�O�p��</option>
<option style="color:#FFFFFF" value="/�ͤ�J�|$ ">�ͤ�J�|</option>
<option style="color:#FFFFFF" value="/�o�g�l�u$ ">�o�g�l�u</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/�s��$ 0">�Ȧ�s�� </option>
<option style="color:#FFFFFF" value="/����$ 0">�Ȧ���� </option>
<option style="color:#FFFFFF" value="/���$ 10000">�Ȧ���� </option>
<option style="color:#FFFFFF" value="/�ǰe�k�O$ 100">�ǰe�k�O</option>
<option style="color:#FFFFFF" value="/�s�k�O$ 100">�s��k�O</option>
<option style="color:#FFFFFF" value="/���k�O$ 100">���^�k�O</option>
<%if info(6)="�x��" or info(6)="�ƴx��" then%>
<option style="color:#FFFFFF" value="/�x���O$ ����������">�x �� �O </option>
<option style="color:#FFFFFF" value="/�P��$ ">��������</option>
<option style="color:#FFFFFF" value="/�Ѱ��P��$ ">�Ѱ��P��</option>
<option style="color:#FFFFFF" value="/�����j��$ ">�����j��</option>
<option style="color:#FFFFFF" value="/�U��$ �����W">�U�ʨ��� </option>
<option style="color:#FFFFFF" value="/�U��$ �Ƥ�">�U�ʰƤ� </option>
<option style="color:#FFFFFF" value="/�U��$ ����">�U�ʪ��� </option>
<option style="color:#FFFFFF" value="/�U��$ �@�k">�U���@�k </option>
<option style="color:#FFFFFF" value="/�U��$ ��D">�U�ʰ�D </option>
<option style="color:#FFFFFF" value="/�U��$ �d��">�U�ʬd�� </option>
<option style="color:#FFFFFF" value="/�U��$ �̤l">�����U�� </option>
<option style="color:#FFFFFF" value="/���W�B�m$ 1000">���W�B�m </option>
<option style="color:#FFFFFF" value="/�ۦ�$ ">�ۦ��̤l</option>
<%end if
if info(6)="����" then%>
</option>
<option style="color:#FFFFFF" value="/�U��$ �@�k">�U���@�k </option>
<option style="color:#FFFFFF" value="/�U��$ ��D">�U�ʰ�D </option>
<option style="color:#FFFFFF" value="/�U��$ �d��">�U�ʬd�� </option>
<option style="color:#FFFFFF" value="/�U��$ �̤l">�����U�� </option>
<option style="color:#FFFFFF" value="/���W�B�m$ 1000">���W�B�m </option>
<option style="color:#FFFFFF" value="/�ץ�$ �мg����] ">�аV�̤l </option>
<option style="color:#FFFFFF" value="/���y$ 5000 ">���y�̤l
<option style="color:#FFFFFF" value="/�ۦ�$ ">�ۦ��̤l
<%end if
if  info(6)="�@�k" then%>
</option>
<option style="color:#FFFFFF" value="/�U��$ ��D">�U�ʰ�D </option>
<option style="color:#FFFFFF" value="/�U��$ �d��">�U�ʬd�� </option>
<option style="color:#FFFFFF" value="/�U��$ �̤l">�����U�� </option>
<option style="color:#FFFFFF" value="/�ץ�$ �мg����] ">�аV�̤l </option>
<option style="color:#FFFFFF" value="/���y$ 100 ">���y�̤l
<option style="color:#FFFFFF" value="/�ۦ�$ ">�ۦ��̤l
<option style="color:#FFFF66">========
<%end if
if  info(6)="��D" then%>
</option>
<option style="color:#FFFFFF" value="/�U��$ �d��">�U�ʬd�� </option>
<option style="color:#FFFFFF" value="/�U��$ �̤l">�����U�� </option>
<option style="color:#FFFFFF" value="/�ۦ�$ ">�ۦ��̤l
<option style="color:#FFFFFF" value="/���y$ 100 ">���y�̤l
<option style="color:#FFFF66">========
<%end if
if info(2)>=6 then%>
</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/��D����$ ">������D </option>
<option style="color:#FFFFFF" value="/�_��$ ">�ϥδ_�� </option>
<option style="color:#FFFFFF" value="/�ץ�$ �мg����] ">�ϥΰץ� </option>
<option style="color:#FFFFFF" value="/ĵ�i$ �мg����] ">�ϥ�ĵ�i </option>
<option style="color:#FFFFFF" value="/�e��$ �мg����] ">�ϥζe��
<% end if
if info(2)>=7 then%>
</option>
<option style="color:#FFFFFF" value="/���c$ �мg����] ">�ϥΧ��c </option>
<option style="color:#FFFFFF" value="/��H$ �мg����] ">�ϥν�H
<% end if
if info(2)>=8 then%>
</option>
<option style="color:#FFFFFF" value="/�I��$ �мg����] ">�ϥ��I�� </option>
<option style="color:#FFFFFF" value="/�ѥ�$ ">�ѥ޾ާ@ </option>
<option style="color:#FFFFFF" value="/�@��$ �мg����]">�ϥλ@�� </option>
<option style="color:#FFFFFF" value="/�ʦD$ �мg����]">�ϥΰʦD </option>
<% end if
if info(2)>=9 then%>
</option>
<option style="color:#FFFFFF" value="/��j$ ">��j </option>
<option style="color:#FFFFFF" value="/�ʸT$ �мg����] ">�ϥκʸT </option>
<option style="color:#FFFFFF" value="/����$ �Τ�W">�Ѱ��ʸT </option>
<option style="color:#FFFFFF" value="/�T���}��$">�T���}�� </option>
<% end if
if info(2)>=10 then%>
</option>
<option style="color:#FFFFFF" value="/���u$">�� �� �u </option>
<option style="color:#FFFFFF" value="/�x�����y$ 10000">�ϥμ��y </option>
<option style="color:#FFFFFF" value="/���i$ ">�x�����i
<option style="color:#FFFFFF" value="/����$">�����}��</option>
<option style="color:#FFFFFF" value="/�dip$ ">�d��IP
<% end if 
if info(2)>=11 then%>
<option style="color:#FFFFFF" value="/���ܨp��">���ܨp�� </option>
<option style="color:#FFFFFF" value="/��������">�������� </option>
<option style="color:#FFFFFF" value="/�d�ݪ��~$">�d�ݪ��~ </option>
<option style="color:#FFFFFF" value="/�������~$ �d��">�����d��
<option style="color:#FFFFFF" value="/����ip$ �мg����]">����IP </option>
<option style="color:#FFFFFF" value="/����ip$">����IP </option>
<% end if %>
</option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>�`�ΩR�O4
<option style="color:#FFFF66">=�o����r=</option>
<option style="color:#FFFFFF" value="/�ť���r$ ">�ť���r </option>
<option style="color:#FFFFFF" value="/�����r$ ">�����r </option>
<option style="color:#FFFFFF" value="/������r$ ">������r </option>
<option style="color:#FFFFFF" value="/������r$ ">������r </option>
<option style="color:#FFFFFF" value="<I>���r ">����r </option>
<option style="color:#FFFFFF" value="<B>���r ">����r </option>
<option style="color:#FFFFFF" value="<BIG>���r ">�j�r�� </option>
<option style="color:#FFFFFF" value="<CENTER>���r ">�V����� </option>
<option style="color:#FFFFFF" value="<u>���r</u> ">���u�r </option>
<option style="color:#FFFFFF" value="<marquee>���r</marquee> ">�����O </option>
<option style="color:#FFFF66">=====</option>
<option style="color:#FFFFFF" value="/�e�|�O$ 10000">�ذe�|�O </option>
<option style="color:#FFFFFF" value="/�e����$ 10000">�ذe���� </option>
<option style="color:#FFFFFF" value="/�M��y�P$ ">�M��y�P </option>
<option style="color:#FFFFFF" value="/�M���s�]$ ">�M���s�] </option>
<option style="color:#FFFFFF" value="/�M��y�P�\$ ">�M��y�P�\ </option>
<option style="color:#FFFF66">=����ƭ�=</opton>
<option style="color:#FFFFFF" value="/�����O$ ">�����O </option>
<option style="color:#FFFFFF" value="/������O$ ">������O</option>
<option style="color:#FFFFFF" value="/����Z�\$ ">����Z�\ </option>
<option style="color:#FFFFFF" value="/�������$ ">������� </option>
<option style="color:#FFFFFF" value="/������s$ ">������s </option>
<option style="color:#FFFFFF" value="/����y�O$ ">����y�O </option>
<option style="color:#FFFFFF" value="/����D�w$ ">����D�w</option>
<option style="color:#FFFFFF" value="/����k�O$ ">����k�O </option>
<option style="color:#FFFFFF" value="/������$ ">������ </option>
<option style="color:#FFFFFF" value="/����|���O$ ">����|�O </option>
<option style="color:#FFFFFF" value="/�������$ ">������� </option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>�`�ΩR�O5
<option style="color:#FFFF66">=�o���r=</option>
<option style="color:#FFFFFF" value="/�ť��r$ ">�ť��r </option>
<option style="color:#FFFFFF" value="/�����r$ ">�����r </option>
<option style="color:#FFFFFF" value="/����r$ ">����r </option>
<option style="color:#FFFFFF" value="/�����r$ ">�����r </option>
<option style="color:#FFFFFF" value="/�����r$ ">�����r </option>
<option style="color:#FFFFFF" value="/�����r$ ">�I���r </option>
<option style="color:#FFFF66">=�s�W�\��=</option>
<option style="color:#FFFFFF" value="/�ť���r$ �hhttp://lok332005.somee.com�A�G���򦳦n�h�\��A�������A����ө��A�`���D���U�G�ɤ��ФH�G�A�A�A�A���U���|���������|��&5000����">�ԤH��d�� </option>
<option style="color:#FFFFFF" value="/�ť���r$ ����t��*�ƭȰө�v2.0*������*�ۧU�[�ިt��*�l�h���d*�S�O�r��*�ǳN*�ǩG*�s�i�C*���|�u�X��*���|�L�k��ܺ���*����!... ">�����\�� </option>
<option style="color:#FFFFFF" value="/�٭��N$ ">�٭��N </option>
<option style="color:#FFFFFF" value="/�հ��ʯv�N$ ">�հ��ʯv�N </option>
<option style="color:#FFFFFF" value="/�I�}�ѯv�N$ ">�I�}�ѯv�N </option>
<option style="color:#FFFFFF" value="/�j��X���N$ ">�j��X���N </option>
<option style="color:#FFFFFF" value="/�a����޳N$ ">�a����޳N </option>
<option style="color:#FFFFFF" value="/�·t���N$ ">�·t���N </option>
<option style="color:#FFFFFF" value="/�·t��N$ ">�·t��N </option>
<option style="color:#FFFFFF" value="/�·t����N$ ">�·t����N </option>
<option style="color:#FFFFFF" value="/�·t����N$ ">�·t����N </option>
<option style="color:#FFFFFF" value="/�·t��|�N$ ">�·t��|�N </option>
<option style="color:#FFFFFF" value="/�a������N$ ">�a������N </option>
<option style="color:#FFFFFF" value="/�����[�޳N$ ">�����[�޳N </option>
<option style="color:#FFFFFF" value="/�����[���N$ ">�����[���N </option>
<option style="color:#FFFFFF" value="/�����[���N$ ">�����[���N </option>
<option style="color:#FFFFFF" value="/�����[��N$ ">�����[��N </option>
<option style="color:#FFFFFF" value="/�����[���N$ ">�����[���N </option>
<option style="color:#FFFFFF" value="/�����[�|�N$ ">�����[�|�N </option>
<option style="color:#FFFFFF" value="/�����[���N$ ">�����[���N </option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>�ʶR�ƭ�
<option style="color:#FFFF66">==�ʶR�ƭ�===</option>
<option style="color:#FFFFFF" value="/�ʶR�|�O$ ">�ʶR�|�O </option>
<option style="color:#FFFFFF" value="/�ʶR����$ ">�ʶR���� </option>
<option style="color:#FFFFFF" value="/�ʶR�l�u$ ">�ʶR�l�u </option>
<option style="color:#FFFFFF" value="/�ʶR���$ ">�ʶR��� </option>
<option style="color:#FFFFFF" value="/�ʶR�k�O$ ">�ʶR�k�O </option>
<option style="color:#FFFF66">==�ʶR��O===</option>
<option style="color:#FFFFFF" value="/�ʶR��O$ ">�ʶR��O </option>
<option style="color:#FFFFFF" value="/�ʶR�Z�\$ ">�ʶR�Z�\ </option>
<option style="color:#FFFFFF" value="/�ʶR����$ ">�ʶR���� </option>
<option style="color:#FFFFFF" value="/�ʶR���s$ ">�ʶR���s </option>
<option style="color:#FFFFFF" value="/�ʶR�y�O$ ">�ʶR�y�O </option>
<option style="color:#FFFFFF" value="/�ʶR�D�w$ ">�ʶR�D�w </option>
<option style="color:#FFFF66">==�ʶR�|��===</option>
<option style="color:#FFFFFF" value="/�ʶR2�ŷ|��$ ">�ʶR2�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR3�ŷ|��$ ">�ʶR3�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR4�ŷ|��$ ">�ʶR4�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR5�ŷ|��$ ">�ʶR5�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR6�ŷ|��$ ">�ʶR6�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR7�ŷ|��$ ">�ʶR7�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR8�ŷ|��$ ">�ʶR8�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR9�ŷ|��$ ">�ʶR9�ŷ|�� </option>
<option style="color:#FFFFFF" value="/�ʶR10�ŷ|��$ ">�ʶR10�ŷ|�� </option>
<option style="color:#FFFF66">==�ʶR�x��===</option>
<option style="color:#FFFFFF" value="/�ʶR6�ũx��$ ">�ʶR6�ũx�� </option>
<option style="color:#FFFFFF" value="/�ʶR7�ũx��$ ">�ʶR7�ũx�� </option>
<option style="color:#FFFFFF" value="/�ʶR8�ũx��$ ">�ʶR8�ũx�� </option>
<option style="color:#FFFFFF" value="/�ʶR9�ũx��$ ">�ʶR9�ũx�� </option>
<option style="color:#FFFFFF" value="/�ʶR10�ũx��$ ">�ʶR10�ũx�� </option>
<option style="color:#FFFF66">==¾�~�ഫ===</option>
<option style="color:#FFFFFF" value="/�ʶR���Ѯv$ ">�ʶR���Ѯv </option>
<option style="color:#FFFFFF" value="/�ʶR���s�Ԥh$ ">�ʶR���s�Ԥh </option>
<option style="color:#FFFFFF" value="/�ʶR����i�h$ ">�ʶR����i�h </option>
</select>
<input type=button value='�_��' onClick="window.open('../yamen/disp.asp','casper','scrollbars=yes,resizable=yes,width=400,height=300')" style="background-color:FF6633;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2">
<input type=button value='�K��' onClick="window.open('pic.asp','f3')" style="background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " name="button">
<input type=button name="tbclutch" value="����" onClick="javascript:parent.tbclutch();" title="�X��/����/���̤W�j/���̤U�j/��������" onMouseOver="window.status='�X��/����/�W�j/�U�j/���������C';return true" onMouseOut="window.status='';return true" style="background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> <td> <div align="center"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> <td height="20">
<div align="center">
<input type="checkbox" name="addvalues">
<a href="#" onClick='document.af.addvalues.checked=!(document.af.addvalues.checked);document.af.sytemp.focus();'>
<font color="#90CFEF">�w�I</font></a>
<input type='checkbox' name='sp' accesskey='v' onClick='parent.f1.speed();document.af.sytemp.focus();' value="ON">
<a href=# onClick='document.af.sp.checked=!(document.af.sp.checked);parent.f1.speed();document.af.sytemp.focus();' onMouseOver="window.status='�y�Z����ܹ�ܰϤ�r�C';return true" onMouseOut="window.status='';return true" title="���ҥ�/�������y�Z�u�̤覡|�ֱ���Alt+V">
<font color="#90CFEF">��</font></a><a href="#"><font color="#90CFEF">��</font></a>
<input onclick='document.af.listname.focus();if(document.af.listname.checked==true){parent.m.location.reload();}'
type=checkbox name=listname title='���H�i�J�ΰh�X���ɭ�,�۰ʨ�s�W��;' value="ON" checked>
<a href='#' onclick='document.af.listname.checked=!(document.af.listname.checked);
document.af.listname.focus();if(document.af.listname.checked==true){parent.m.location.reload();}' title='���H�i�J�ΰh�X���ɭ�,�۰ʨ�s�W��; '>
<font color="#90CFEF">�W��</font></a>
<input onClick='document.af.soundtx.focus();' type=checkbox name=soundtx checked title='�O�_�ϥβ�ѫ��n���B���ֵ��I' value="ON">
<a href='#' onClick='document.af.soundtx.checked=!(document.af.soundtx.checked);document.af.soundtx.focus();' title='�O�_�ϥβ�ѫ��n���B���ֵ��I'>
<font color="#90CFEF">��</font></a><a href="#"><font color="#90CFEF">��</font></a>
<input type=checkbox name='towhoway' value='1' accesskey='s' onClick="document.af.sytemp.focus();">
<a href='#' onMouseOver="window.status='�襤������A�����ܥu���A�M���~��ݨ�]�Y�ϯ����]�ݤ���^�C';return true" onMouseOut="window.status='';return true" onClick="document.af.towhoway.checked=!(document.af.towhoway.checked);document.af.sytemp.focus();" title="�����ܨந����">
<font color="#90CFEF">�p��</font></a>
<input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON">
<a href='#' onclick='document.af.as.checked=!(document.af.as.checked);
document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='�O�_�u�����|�ֱ���Alt+A '>
<font color="#90CFEF">�u��</font></a>
<input type="checkbox" name="titleline" value="1" accesskey='t' onClick="document.af.sytemp.focus();">
<a href='#' onMouseOver="window.status='�N�o�����e�m���ѫǪ��Τ���D��C';return true" onMouseOut="window.status='';return true" onClick="document.af.titleline.checked=!(document.af.titleline.checked);document.af.sytemp.focus();" title="�G�ťH�W��ͤ~��ϥ�">
<font color="#90CFEF">���D</font></a>
<%if info(2)>=6  then%>
<a href="../dt/ask.asp" target="_blank"><font color="#90CFEF">�o�D</font></a>
<a href="../tongji/tj.asp" target="_blank"><font color="#90CFEF">�q��</font></a>
<a class=blue href="#" onClick="window.open('../dalie/zj.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="�x�����y���F�I">
<font color="#90CFEF">�y��</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fyuanbao.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�x�����_�F�I">
<font color="#90CFEF">���_</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fjiangshi.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�x���񧯩ǤF�I">
<font color="#90CFEF">����</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fdog.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�x���񪯪��F�I">
<font color="#90CFEF">�p��</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/dabianfu.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�x����ǳ��F�I">
<font color="#90CFEF">�ǳ�</font></a>
<a class=blue href="#" onClick="window.open('../jiaofei/zj.asp','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�g��X�{�F�I">
<font color="#90CFEF">�g��</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fq.asp','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="�o�F��I">
<font color="#90CFEF">�o�F��</font></a>
<%else%>
<a href="../dt/reply.asp" target="_blank"><font color="#90CFEF">���D</font></a>
<a class=blue href="#" onClick="window.open('../dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="�u�n�x�����y���F�A�A�i�H���y���o���I">
<font color="#90CFEF">���y</font></a>
<a class=blue href="#" onClick="window.open('../jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="�u�n���g��F�A�A�i�H���g��o�F��F�I">
<font color="#90CFEF">���g��</font></a>
<%end if%>
<%if info(2)>=8 then%>
<a href="guanadmin/fine.asp" target="_blank"><font color="#90CFEF">��H</font></a>
<%end if%>
<input type="hidden" name='towhoid' readonly value='0' size=2 maxlength=10>
<%if info(2)>=10 then%>
<a href="../linjianww/login.asp" target="_blank"><font color="#FF0000">����</font>
<%end if%>
</a>
</div>
</td></tr>
</table><script language="JavaScript">function startnosay(){var nosay=99999;document.af.clock.value=nosay;}</script> 
<script src="readonly/f2.js"></script> <%if id=1 then%> <script>parent.rn();</script> 
<%else%> 
<script>parent.fc();</script> 
<%end if%>
<div id="dh" style="position:absolute; left:-80px; top:-80px; width:0px; height:0px; z-index:1"> 
<input type="button" value="<" name='gopp' onClick="Javascript:gop();" accesskey=","> 
<input type="button" value=">" name='gonn' onClick="Javascript:gon();" accesskey="."> 
</div></div></td></tr> </table></form></div>
</body></html>
