<!--#include file="config.asp"-->
<!--#include file="conn.asp"-->


<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5"> 
<head>
<style type="text/css">
<!--
a {  text-decoration: none}  
a:hover {  text-decoration: underline} 
table {  font-size: 9pt}
body,table,p,td,input,select,textarea {  font-size: 9pt} 
-->
</style>
<title>k666�n��� ���� http://www.vv66.net</title>
</head>

<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<form method="post" action="save.asp" align="center">
  <center><table
    border="1" cellspacing="1" width="400" bordercolor="<%=titlelightcolor%>"
    height="310">
        <tr>
            
        <td align="center" colspan="2" bgcolor="<%=titledarkcolor%>"><font color="#FFFF00">�i<%=title%>�j 
          ��D�@��</font></td>
        </tr>
        <tr>
            <td colspan="2" height="23"><p align="center">&nbsp;<font
            color="<%=titlelightcolor%>">�a �]*�^ ������g </font></p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">�z���m�W�G </p>
            </td>
            <td height="23"><p align="left"><input type="text"
            size="20" maxlength="20" name="name"> �]*�^ </p>
            </td>
        </tr>

        <tr>
            <td align="center" width="100" height="24"><p
            align="left">�z���ʧO�G </p>
            </td>
            <td height="24"><p align="left"><select name="sex"
            size="1">
                <option value="�k"> �k </option>
                <option value="�k"> �k </option>
            </select> </p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">�z���H�c�G </p>
            </td>
            <td height="23"><p align="left"><input type="text"
            size="30" maxlength="50" name="email"> �]*�^ </p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">�z���D���G </p>
            </td>
            <td height="23"><p align="left">
            <input type="text"
            size="30" maxlength="50" name="homepage"
            value="http://">
             </p>
            </td>
        </tr>
        <tr>
            <td width="100" height="23"><p align="left">�@�����O�G
            </p>
            </td>
            <td height="23"><p align="left">
            <select name="wishtype"
            size="1">
              <option value="love">��  �R</option>
                <option value="study">��  �~</option>
                <option value="health">��  �d</option>
                <option value="family">�a  �x</option>
                <option value="work">��  �~</option>
                <option value="future">�N  ��</option>
                <option value="wealth">�]  �I</option>
                <option value="life">��  ��</option>
            </select> �]*�^ </p>
            </td>
        </tr>
        <tr>
            <td width="100" class="pt9">�~��a��G</td>
            <td class="pt9"><table border="0" width="300">
                <tr>
                    <td>
                <input type="radio" name="address" value="����" checked>����</td>
              <td>
                <input type="radio" name="address" value="9��">9��</td>
                    <td><input type="radio" name="address" value="�饻">�饻 </td>
                    <td><input type="radio" name="address" value="�O�W">�O�W</td>
                    <td><input type="radio" name="address" value="�s�[��">�s�[��</td>
                </tr>
                <tr>
                    <td><input type="radio" name="address" value="���Ӧ��">���Ӧ��</td>
                    <td><input type="radio" name="address" value="�[���j">�[���j</td>
                    <td><input type="radio" name="address" value="�D�w">�D�w</td>
                    <td><input type="radio" name="address" value="����">����</td>
                    <td><input type="radio" name="address" value="�^��">�^��</td>
                </tr>
                <tr>
                    <td><input type="radio" name="address" value="�k��">�k��</td>
                    <td><input type="radio" name="address" value="�D�w">�D�w</td>
                    <td><input type="radio" name="address" value="�D��">�D��</td>
                    <td><input type="radio" name="address" value="�䥦">�䥦</td>
                    <td> �]*�^ </td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td width="100">��D�@��G </td>
            <td>
          <div align="center">
            <textarea name="info" rows="4" cols="45"></textarea>
            <br>
            �]* 100�r�H���^ </div>
        </td>
        </tr>
        <tr>
            
        <td colspan="2" bgcolor="<%=titledarkcolor%>" height="15" align=center><font color="#FFFF00">�e�X��X�_�����ۤ߬�D...</font></td>
        </tr>
        <tr>
            <td align="center" colspan="2" height="25"><input
            type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="�e  �X">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" type="reset" value="�M  ��"> </td>
        </tr>
    </table>
    </center></div>
</form>
<!--#include file="copyright.asp"-->
</body>
</html>
