<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>�ק�K�X</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body background="../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<center> <table border=1 align=center width=326 cellpadding="5" cellspacing="8" background="../images/12.jpg" HEIGHT="350"> 
<tr bgcolor="#FFFFFF" align="center"> <td height="209" background="../ljimage/downbg.jpg"> 
<p>�� �� �K �X</p><table border="0" width="288" height="257"> <tr> <td> <form method=POST action='modifyok.asp'> 
<div align="center">�m �W�G <input type=text name=name size=12 maxlength="10"> <br> 
��K�X�G <input type=password name=oldpass size=12 maxlength="10"> <br> �s�K�X�G <input type=password name=pass size=12 maxlength="10"> 
<br> �T �{�G <input type=password name=repass size=12 maxlength="10"> <br> �{ ���G <input type=text name=regjm1 size=5 maxlength="5" value=<%=regjm%>> 
<br> ������ҳ��п�J�{�ҡG<font color="#FF0000"><%=regjm%></font> <br> <p> <input type=submit value=�ק� name="submit"> 
<input type=button value=���� onClick="window.close()" name="button"> <input type=hidden name=regjm value="<%=regjm%>"> 
</p></div></form></td></tr> <tr> <td style='color:red;font-size:9pt'> <font color="#FF6600">�����G<br> 
<font color="#9900CC">(�K�X���פ��o�W�L10��^</font><br> �ѩ�IE5�|�O����J���K�X�A<br> �Цb���a�W�����B�ͳq�L���M�����v�O����<br> 
�ӧR���O���I�H�K�b���Q�s�� </font></td></tr> </table></table></center>
</body>
</html>
 