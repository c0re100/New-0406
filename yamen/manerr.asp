<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>�X������</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body background=../images/8.jpg leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
        <table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" background="../images/7.jpg">
          <tr>
<td>
<table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="360" align="center"><font color="#FFFFFF" size="3">�X������</font></td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="328" valign="top">
<%Select Case id
Case "000"%>
<p>  �{�ǩҦb�ؿ����O�����ؿ��A�γ]�m�����~�AGlobal.asa �S���Q����C���{�ǻݭn�����ؿ�������I</p>
<%Case "001"%>
<p>  <font color=red><%=Request.QueryString("kickname")%></font> �ä��b��ѫǤ��A�۱ϵ{�ǵL�k�����A�d�ݬO�_�ж���ܿ��~�I</p>
<%Case "002"%>
<p>  �O���O�Q�}���F�H�s�j�W�M�f�O�������H�䦺�r�I</p>
<%Case "003"%>
<p>  �O�ڤ���A�٬O�A����H�n�䪺�H���̦��r�H�³J�I</p>
<%End Select%>
</td>
</tr>
<tr>
<td align="center" valign="top">
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
<html><script language="JavaScript">                                                                  </script></html> 