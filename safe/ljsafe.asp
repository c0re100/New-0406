<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5> 
<title>�ק�K�X</title>
<style type="text/css">
</style>
<link rel="stylesheet" href="Setup.css">
</head>
<body oncontextmenu=self.event.returnValue=false text="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" bgcolor="#990000">
<%
if not IsArray(Session("info")) then 
%>
<div align="center">�藍�_�A�z�S���n���A�ӽбK�X�O�@���ݵn�����i�ӽСI <br> <A HREF="#" ONCLICK="window.close()">[ 
�� �� �� �f ]</A><%
Response.End 
end if
info=Session("info")
%> </div><center> <p><font color="yellow">�� �ӽбK�X�O�@ ��</font></p><table border=1 bgcolor="#008080" align=center cellpadding="10" cellspacing="13" bordercolordark="#FFFFFF" bordercolorlight="#000000" width=380> 
<tr><td bgcolor=#008080 width="378"> <table bgcolor="#008080" border="1" align="center" width="100%" height="10" cellspacing="0" bordercolordark="#008080" bordercolorlight="#008080"> 
<tr> <td height="1"> <form method=POST action='ljsafeok.asp'> <div align="center"> 
<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center"> 
<tr> <td>�m  �W�G</td><td> <input type=text name=name readonly size=13 maxlength="10" value=<%=info(0)%>> 
</td><td>�n�����򪺽㸹</td></tr> <tr> <td>�� �K �X�G</td><td> <input type=password name=oldpass size=13 maxlength="10"> 
</td><td>�n�����򪺱K�X</td></tr> <tr> <td>��^�K�X�G </td><td> <input type=password name=getpass size=13 maxlength="10"> 
</td><td>�n�ӽЪ��K�X�O�@</td></tr> <tr> <td>����K�X�G</td><td> <input type=password name=repass size=13 maxlength="10"> 
</td><td>�n�ӽЪ��K�X�O�@</td></tr> <tr> <td>�l  �c�G</td><td> <input type=text name=mail size=13 maxlength="30"> 
</td><td>�ۤv�`�ζl�c</td></tr> </table><br> </div><div align="center"> <input type=submit value="�� ��" name="submit"> 
<input type=button value="�� ��" onClick="window.close()" name="button"> </div></form><p> 
�ӽЪ��K�X�O�@�i�Ψӥᥢ�K�X���l�ƱK�X�I<br> �]���I�Ф@�w�n�O��z����^�K�X�M�l�c�I </p></td></tr> </table></table></center>  
</body>  
</html> 
 
 
