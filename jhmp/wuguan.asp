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
<title>�F�C����m�Z</title>
<body leftmargin="0" topmargin="0" bgcolor="#003399" background="../linjianww/0064.jpg">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="150" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=info(0)%></font>�w����{�ߪZ��</b></font></div>
<form method=POST action='wuguanok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
            <td align=center><font color=blue>�v��:<%=info(7)%></font>
            <%if trim(info(7))<>"�L" then%>
              <select name=money size=1>
                <option value="1000" selected> �����B�]�@�d��^</option>
                <option value="10000">�֪L�w�\�]�@�U��^</option>
                <option value="100000">�o�ܤߪk�]�Q�U��^</option>
                <option value="1000000">�p�����C�]�@�ʸU��^</option>
                <option value="2000000">�]�a�۬r�]�G�ʸU��^</option>
                <option value="10000000">�E�M����]�@�d�U��^</option>
              </select>
              <%end if%>
</td>
</tr>
<tr>
<td  align=center>
            <%if trim(info(7))<>"�L" then%>
              <input type=submit value=�}�l�m�\ name="submit">
              <%end if%>
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                �ާ@²��</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
            <%if trim(info(7))="�L" then%><div align="center">
            <font color=red>�S���v�Ť���ާ@�I</font><BR></div>
            <%end if%>�v�Ż�A�h�m�\�A��ܧA�A�X�ۤv���ߪk(�����O���P����I)�ץi�H�פѨ�ܰ��L�����Z�\�I</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>



