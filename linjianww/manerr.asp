<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("hxf_c_chatroombgimage")
chatroombgcolor=Application("hxf_c_chatroombgcolor")
nl=""
Select Case id
Case "000"
 nl="�{�ǩҦb�ؿ����O�����ؿ��A�γ]�m�����~�AGlobal.asa �S���Q����C���{�ǻݭn�����ؿ�������I"
Case "100"
 nl="�藍�_�A�A�|���n���Τw�g�W���_�}�s���A�Ц^��n���������s��J�Τ�W�M�K�X�i��n���C"
Case "101"
 nl="�����֦� 3 �ťH�W�v���~��K�[�۳Ъ��ʧ@�C"
Case "102"
 nl="�U�Ӷ��ا������ťաA�ж�g����I"
Case "103"
 nl="�ʧ@���e�����O�H��//���}�Y���y�y�I"
Case "104"
 nl="�ʧ@���e����X�{��//##���A�o�N�ɭP�o����m�W�X�{���_�A�бN ## �h���I"
Case "105"
 nl="�ʧ@�W�٩ΰʧ@���e�W�L������סI"
Case "106"
 nl="�ʧ@���e�t����%%���A�ʧ@�������אּ��1 ��O�H���I"
Case "107"
 nl="�ʧ@��������1 ��O�H���A�ʧ@���e���o�S���X�{��%%���I"
Case "108"
 nl="�K�[���ʧ@���e���A����]�t�b������\���B�������B��" & chr(34) & "���I"
Case "109"
 nl="�Ӥ��e�w�g�s�b�A�������_�K�[�C"
Case "120"
 nl="�A�S���M�z����Ȥ��}�����v���C"
Case "121"
 nl="�S���W�L 7 �Ѫ��O���A����M���C"
Case "130"
 nl="�A�S�����b���޲z�����v���C"
Case "131"
 nl="�S�������b���i�ѧR���C"
Case "132"
 nl="�п�J�Τ�W�C"
Case "133"
 nl="�b�w���۱������b�����S�����ӥΤ�W�A�����_�C"
Case "134"
 nl="�����_�Τ�W�G<font color=red>" & Request.QueryString("name") & "</font>�A�]���t�Τ��w���ۦP���Τ�W�s�b�C�p�G�A�T��Q�n��_�ӥΤ�W�A�h�ХΡ��R���b�����\��A���R���t�Τ��ۦP���Τ�W��A��_�ӥΤ�W�C"
Case "135"
 nl="�п�J���R�����Τ�W�C"
Case "136"
 nl="�Τ�W���s�b�A����R���C"
Case "137"
 nl="�Τ�W���s�b�C"
Case "138"
 nl="�ӥΤ�W���w�ä[�O�d�A�������_�ާ@�C"
Case "139"
 nl="�ӥΤ�W���Q�ä[�O�d�A��������C"
Case "140"
 nl="�¥Τ�W�M�s�Τ�W�����ର�šC"
Case "141"
 nl="����N�Τ�W�אּ�G<font color=red>" & Request.QueryString("name") & "</font>�A�]���t�Τ��w���ۦP���Τ�W�s�b�C�p�G�A�T��Q�n�令�o�ӥΤ�W�A�h�ХΡ��R���b�����\��A���R���t�Τ��ۦP���Τ�W�C"
Case "142"
 nl="�Τ�W�M�s�K�X�����ର�šC"
Case "150"
 nl="�A�S�����ƾ����Y�����v���C"
Case "151"
 nl="�ƾڮw�|�������A�������s���}�ƾڮw�C"
Case "152"
 nl="�ƾڮw�|�����}�A���������ƾڮw�C"
Case "160"
 nl="�п�J�n�j�����Τ�W�C"
Case "200"
 nl="�A�S������H�����v���C"
Case "201"
 nl="�A���b��ѫǤ��A������桧��H���ާ@�C"
Case "202"
 nl="�Ы��w�n��X���ﹳ�C"
Case "203"
 nl="���o�L�G��H�A�п�J��]�C"
Case "204"
 nl="�Τ�W�G<font color=red>" & Request.QueryString("kickname") & "</font> ���b��ѫǤ��A�����A��F�C"
Case "205"
 nl="�����A��f�A��ۤv���ڡC"
Case "210"
 nl="�A�S����IP�޲z�����v���C"
Case "211"
 nl="�A���b��ѫǤ��A����i�桧IP�޲z���C"
Case "212"
 nl="�Ы��w�n���ꪺ IP�C"
Case "213"
 nl="����ۤv�� IP�H�O�̤F�I"
Case "214"
 nl="���o�L�G���� IP�A�п�J��]�C"
Case "215"
 nl="�� IP �w�g�Q����F�A���୫�_����C"
Case "216"
 nl="�п�J����� IP ����]�C"
Case "217"
 nl="�� IP ���Q����A����i�����C"
Case "218"
 nl="�Ы��w�n���ꪺ IP�C"
Case "219"
 nl="�n���ꪺIP�P�Τ�W�������C"
Case "220"
 nl="�A�S�������u�ާ@�����v���C"
Case "221"
 nl="�A���b��ѫǤ��A����i�桧���u�ާ@���C"
Case "222"
 nl="�Ы��w�n�F�����ﹳ�C"
Case "223"
 nl="�ڡA�A���O�Z�Ҹ̥��OŢ�V�V�䦺�a�C"
Case "224"
 nl="���o�L�G�å����u�A�п�J��]�C"
Case "225"
 nl="�Τ�W�G<font color=red>" & Request.QueryString("bombname") & "</font> ���b��ѫǤ��A������F�C"
Case "230"
 nl="�A�S�����t�ΰѼơ����v���C"
Case "231"
 nl="�s�ȻP�­ȧ����ۦP�A�����i����C"
Case "232"
 nl="��J���Ȥ��X�k�C"
Case "240"
 nl="�A�S���i�桧�ŧO�޲z�����v���C"
Case "241"
 nl="�A���b��ѫǤ��A����i�桧�ŧO�޲z���C"
Case "242"
 nl="�A�S�����桧�ɯžާ@�����v���C"
Case "243"
 nl="�A�S�����桧���žާ@�����v���C"
Case "244"
 nl="�Τ�W���ର�šC"
Case "245"
 nl="�䤣��Τ�W�G<font color=red>" & Request.QueryString("username") & "</font>�C"
Case "246"
 nl="�ӥΤ�W���ŧO����A�C�A������i��ާ@�C"
Case "247"
 nl="��w���ŧO�Ȥ��X�k�C"
Case "248"
 nl="�п�J��]�C"
Case "250"
 nl="�A�S���d�ݡ�HTML���Ρ����v���C"
Case "255"
 nl="�A�S�����������i�����v���C"
Case "256"
 nl="���e���ର�šC"
Case "260"
 nl="�A�S�����ʧ@�޲z�����v���C"
Case "261"
 nl="�䤣��Ӱʧ@�C"
Case "270"
 nl="�A�S�����d���޲z�����v���C"
Case "271"
 nl="�ӯd�����s�b�C"
Case "280"
 nl="�A�S�����վ�ŧO�����v���C"
Case "281"
 nl="�A���b��ѫǤ����ࡧ�վ�ŧO���C"
Case "282"
 nl="�Τ�W���ର�šC"
Case "283"
 nl="�п�J�վ�ŧO����]�C"
Case "300"
 nl="�A�S���޲z���벼�t�Ρ����v���C"
Case "301"
 nl="���ŦX�벼����A����벼�C"
Case "302"
 nl="�榡���~�C"
Case "303"
 nl="�п�ܧA������Կ�H�C"
Case "304"
 nl="�Կ�H���s�b�C"
Case "305"
 nl="�Կ�H���ର�šC"
Case "306"
 nl="�Կ�H�w�g�s�b�A���୫�_�K�[�C"
Case "310"
 nl="�A�S�����ä[���ꡨ�עު��v���C"
Case "311"
 nl="�榡���~�C"
Case "320"
 nl="IP���ର�šC"
Case else
 nl="�藍�_�A�ӥX���������Q�n�O�C"
End Select
nl="<p>.2' g Sl[dBu<A\>'%><html>
<head>
<title>�X������</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center"> 
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
              <table border="0" bgcolor="#3399FF" cellspacing="0" cellpadding="2" width="350">
                <tr>
<td width="342"><font color="FFFFFF"> �X������</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="����"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr> 
                  <td width="59" align="center" valign="top"><font face="Wingdings" color="#0066FF" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value=" �T �w " onclick=javascript:history.go(-1)>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>
 