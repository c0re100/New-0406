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
<title>�������</title>
<body bgcolor="#660000" background="../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<table width="97" border="0" align="center" cellpadding="0" cellspacing="0" background="../chat/bk.jpg">
  <tr> 
<td height="81" valign="top"> 
      <div align="center">�w��<font color="#000000"><font color=#FFFFCC><%=info(0)%></font></font>���{��������w</div>
      <form method=POST action='putmpjj.asp'> 
<table width="300" align="center"> <tr> <td> <tr> <td align=center>�п�ܩҥH�s�J�����ơG 
<select name=money size=1 style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
                <option value="1000" selected>1000</option>
                <option value="10000">10000</option>
                <option value="100000">100000</option>
                <option value="1000000">1000000</option>
                <option value="10000000">10000000</option>
              </select> </td></tr> <tr> <td  align=center> <input type=submit value=�n�F�N�o�� name="submit"> 
</td></tr> <tr> <td valign="top" height="8" > <div align="center"><br> <br> �������</div></td></tr> 
<tr> <td valign="top" > <p>��������A�O�@����������A�Τ�������F���A�Ҩ��o���B�s��󥻪������w���A�Ѱ��k���Ѥ��t���һݭn���H�C�z�糧�����I�X�ڷ|���������^�����C��󥻪�����̤j�αq���������̦h���H�ڭ̳��|���O���Ʀ檺�A�Ш즿��Ʀ椤�d�ߡI</p></td></tr> 
</table></form></td></tr> </table>
</body>
</html>
