<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><script>
if(!confirm("�A���o�ӥ��ƶܡH"))
window.location=window.close()
</script>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>�Ĥ@��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body bgcolor="#0099CC" oncontextmenu=self.event.returnValue=false background="images/bg.gif" LEFTMARGIN="0" TOPMARGIN="0" >
<table width="562" border="0" cellspacing="0" cellpadding="0" align="center"> 
<tr> <td width="595"> <div align="left"><b>�_�ê��Ĥ@���O�W�t�D�Ѫ����͡V�V�ѷ��i�ѹD</b></div></td><td width="162" rowspan="2" valign="top"><img src="jh1.jpg" width="160" height="250"></td></tr> 
<tr> <td width="595" valign="top"> <p style="line-height: 200%; margin: 20"> <font color="#FF3366">�i�ѹD�G</font>�n�p�l�A�o�̬O�_�ê��Ĥ@���A�ڦѤH�a�o��~���N����A�ʤ�F�A���̤U�L�ѡA�AĹ�F�ڴN��A�L�h�C<br><br> 
<font color=red><%=info(0)%></font>:<a href="go1.asp">�n�a�A�ڤ@�w�|�UĹ�A���C</a><br> </p></td></tr> 
</table>
</body>
</html>