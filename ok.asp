<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
Select Case id
Case "100"
nl="���\�ק�K�X�A�s�K�X�� <font color=red>" & Request.QueryString("new") & "</font>�A�аO�c�C"
Case "101"
nl="�۱��ާ@�����I�Τ�W <font color=red>" & Request.QueryString("name") & "</font> �w�g�Q�R���A�ӥΤ�W���Ҧ��O�����v�����w�����C"
Case "110"
nl="�ӤH�H���w�g�ק令�\�C"
Case "111"
nl="�A��W�ާ@�����I"
Case "200"
nl="�����w�g���\�o�e��<font color=red>" & Request.QueryString("name") & "</font>�C�i�H�b�����C���d�ݸӮ����O�_�w�g�Q���X�C"
Case "210"
nl="�ʧ@�w�g�K�[�����C"
Case "300"
nl="�ƾڮw�ƥ������A�Ш� backup �ؿ��U�� global.mdb �i�����Y�C"
Case "301"
nl="�w�g�����Y�᪺�ƾڮw�л\�¼ƾڮw�A�T�{�ҽT��A�ЧR���ӳƥ��ƾڮw�A�H���Q�L�H�U���C"
Case "302"
nl="�ƥ��ƾڮw�R�������C"
Case "600"
nl="���B�n�O���\�I�����O��1000��I"
Case "601"
nl="�g�L�������~�D�A�A�o�{�ۤv�M�n�F���֡A�믫�]�n�h�F�A�]�\�Ө�a�̺Τ@ı�|��[�}�ߧa�C"
Case "700"
nl="���ߧA�I�ƾڮw��s���\�I"
Case "701"
nl="������Ƥw���\�ק�I"
Case "702"
nl="�M�Υ]�]�m�����I�ЧR�� setup.asp ���A�H�K���_�B��I"
Case "703"
nl="�x���Q�}���F�I"
Case "704"
nl="�R�����\�I�I"
Case "705"
nl="����i�ܡG�x�űϨa�u�@���Q���u�I"
Case "706"
nl="����i�ܡG�x�Ŧ�ú�ʩm�ľ����u�@���Q���u�I"
Case "707"
nl=""
Case else
nl="�藍�_�A���������Q�n�O�C"
End Select
nl="<p>  " & nl & "</p>"%><html>
<head>
<title>�ާ@���\</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false  background=JHIMG/Bk_hc3w.gif leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#dcd8d0">
<tr>
<td>
<table border="0" cellspacing="0" cellpadding="2" width="370" background="jhimg/title.jpg">
<tr>
<td width="344"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">�ާ@���\</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="22">
<tr>
<td width="18"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="��^"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="348" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="267"> <%=nl%> </td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<%if id="200" then
Response.Write "<input type='button' name='ok' value=' �d�ݦC�� ' onclick=javascript:top.location.href='webicqlist.asp'>"
else
Response.Write "<input type='button' name='ok' value=' �� �^ ' onclick='javascript:history.go(-1)'>"
end if%>
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