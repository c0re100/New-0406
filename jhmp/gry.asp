<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>�t��|</title>
<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center"> <tr> 
<td height="81" valign="top"> 
      <div align="center"><b><%=info(0)%>���{�F�C�t��|</b></div>
      <form method=POST action='gryok.asp'> 
<table width="300" align="center"> <tr> <td> <tr> <td align=center>�ڪ��R�ߡG <select name=money size=1> 
<option value="1000" selected>1000</option> <option value="10000">10000</option> 
<option value="100000">100000</option> <option value="1000000">1000000 <option value="10000000">10000000</option> 
</select> </td></tr> <tr> <td  align=center> <input type=submit value=���ֳo�ǫĤl name="submit"> 
</td></tr> <tr> <td valign="top" height="8" > <div align="center"><br> <br> �t��|²��</div></td></tr> 
<tr> <td valign="top" > <p>���򪺾Ԫ��B�h�a�ϱo�L�̳o�ǫĤl���h�F�����ˡA�L�U�L���L�̦h��Ʊ�o��ڭ̳̯u�۪����U�A���m�X�A���R�ߡA���@�ɩҪ��t�஽�m�@���R�ߡI(���m�R�ߥi�H�ϱz�w���河��1�I�D�w=500��I)</p></td></tr> 
</table></form></td></tr> </table>
</body>
</html>



