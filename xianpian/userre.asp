<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!-- #include file="conn.asp" -->
<%
Session.TimeOut=30
'�P�_�Τ��檺���
Dim username,username_mail,password,list,strSQL
If request.form("username") = "" Then
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��ƭק�</title>
<style>
<!--
td           { font-family: �s�ө���; font-size: 9pt; color: #000000 }
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style></head>
<body topmargin="0" leftmargin="0" background="../ljimage/11.jpg">
<p align="center"><p align="center"><p><a href=./>��^��ͦC��</a> </p>
<form method="POST" action="userre.asp"><div align="center"><center>  <table border="1" width="600" cellspacing="0" bordercolor="#988CD8">
<tr><td colspan="5" bgcolor="#988CD0">
<p align="center"><font color="#FFFF80" size="3"><b>�� �� �� �� �H ��</b></font></p> 
</td> </tr> </center> <tr> <td>
          <p align="right">�A���m�W�G</p>
</td><center>  <td><p align="center"><input maxLength="20" name="username" size="15" tabIndex="1"></p>
</td></center><td><p align="right">�A���K�X�G</p></td><center>  
<td><p align="center"><font color="#0080ff"><input maxLength="20" name="password" size="20" tabIndex="2" type="password" value></font></p>
</td><td><p align="center"><input tabIndex="10" type="submit" value="   �i�J   "></p>
</td></tr></table></center></div></form><p> </p><p> </p><p> </p>
<form method="POST" action="userre.asp"><div align="center">
<center><table border="1" width="600" cellspacing="0" bordercolor="#988CD8"><tr><td colspan="4" bgcolor="#988CD0"><p align="center"><font color="#FFFF80" size="3"><b>�K �X �d ��</b></font></td> 
</tr> <tr> <td> 
            <p align="center">�A���m�W�G</p>
          </td><td><p align="center"><input maxLength="20" name="username_mail" size="14" tabIndex="3"></p></td>
<td><p align="center">�A���K�X�|�۰ʵo�e��A�ӽЮɨϥΪ��H�c��</p></td><td><p align="center"><input tabIndex="10" type="submit" value="   �T�w   "></p></td>
</tr></table></center></div></form><center></center>    <p> </p><p> </p>
<p> 
  <center>
    <font style="color: #000000; font-family: �s�ө���; font-size: 9pt"> Copyright &copy; 
    2001-2003 cidu.net All rights reserved<br>
    ���v�Ҧ��G<a href="http://xajh.us" target="_blank">�ڤۤu�@��</a>  �������@�G<a href="mailto:seven.s-888@yahoo.com.tw" target="_blank">�ڤ۱��t</a> 
    </font> 
  </center>
</p>
</body>   </html>      

<%Else
username=Request.form("username")
password=request.form("password")
If request.form("username_mail") <> "" Then
'�o�l�󵹥Τ�A���ݭn���˥Τ᪺�W�٩MMAIL�A�M��N�K�X�o�e���Τ�

End If
'�˴��Τ�W�٩M�K�X
If request.form("username")="" or request.form("password")="" Then
 response.write "<script language='javascript'>" & chr(13)
Response.Write "���~�I�Э��s���T��J�I"
 response.write "window.document.location.href='userre.asp';"&Chr(13)
 response.write "</script>" & Chr(13)
Response.End
Else	
		strSQL="select * from list where name='"+Request.form("username")+"'"
		'���˥Τ᪺�W�١A�M���ӱK�X�O�_���T
		  set list=server.createobject("adodb.recordset")

    list.open strSQL,conn,adOpenKeySet,adLockPessimistic
         if list.eof and list.bof then
         %>
            <div align="center"><center>
            <table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: �s�ө���; font-size: 9pt; color: #000000">
            <tr><td width="100%">
            <p align="center"><font color="#FF0000">
            Sorry! �S���A�Q�䪺��ơI</font></p> 
            </td></tr></table></center></div> 
          <%
        response.write "<script language='javascript'>" & chr(13)
        response.write "alert('�S���ӥΤ�I�Ъ`�U.');" & Chr(13)
        response.write "window.document.location.href='add.asp';"&Chr(13)
        response.write "</script>" & Chr(13)
          response.end
          ENd If


        If request.form("password")<>list("password") or list("password")="" or request.form("password")="" Then
        Response.Write "�K�X���~�I<a href = javascript:history.back()>��^����</a>"
        Else
        Session("username") =list("name")
        Session("password") =list("password")
        response.write "<script language='javascript'>" & chr(13)
        response.write "window.document.location.href='moduser.asp?UserID="&list("id")&"';"&Chr(13)
        response.write "</script>" & Chr(13)
        End If
list.close
End if
ENd if
%>