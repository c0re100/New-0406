<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp" -->
<%
dim copyright
dim Aff,I
dim list,MaxPage,EditUserID,xiu
Aff = Chr(34)
Dim strSQL,Keys,UserID,Action
copyright = "<center><font style=" & Aff & "color: #000000; font-family: �s�ө���; font-size: 9pt" & Aff & ">" & vbCrLf
copyright = copyright & "Copyright &copy; 2000-2003  " & vbCrLf
copyright = copyright & " All rights reserved<br>" & vbCrLf
copyright = copyright & "���v�Ҧ��G<a href=" & Aff & "http://21ex.com" & Aff & " target=" & Aff & "_blank" & Aff & ">�ڤۤu�@��</a>" & vbCrLf
copyright = copyright & "�������@�G<a href=" & Aff & "mailto:21ex@21cn.com" & Aff & "></a>"  & vbCrLf
copyright = copyright & "</font></center>" & vbCrLf
'�{�Ƕ}�l�A���T�ӧP�_
'1�B�P�_�W�ťΤ�b���O�_����
'2�B�P�_�O�_���W�ťΤ�
'3�B�i��Τ�޲z
'4�B�`�P�W�ťΤᨭ��
'�����d�ݪ�O�_���޲z�������b���A�p�G�S���h�n�D�`�U�C
'�p�G���޲z����ƫh�M���檺��檺�����i������
'�p�G���ҳq�L�Y�i�H�i�J�޲z�ɭ�
'�p�G���q�L�h��X���~�H���A����^index����
'�p�G�h�X�޲z���ҫh�`�P��e�Τ�C
'//showpage_admin_login             '//�l�{�ǥΨ���ܺ޲z���n������
'//ShowRegAdminPage                 '//�l�{�ǥΨ���ܪ`�U�޲z������
'//SQL_Server                       '//�ƾڮw�ާ@
'//DeleteOrEditOfAdmin              '//�l�{�ǥΨ���ܺ޲z�Τ᭶��
'//chk_admin_password               '//�˴��n���Τ᪺�K�X�O�_���T
'//Write_admin_data                 '//�g�J�ƾڮw�W�ťΤ᪺�H��
'//delete_user_list                 '//�R���Τ�
'//CHK_admin_ISNULL                 '//�ˬd�ƾڪ�O�_���Ŧr�q
'------------------Program Start --------------------------------------
'�ثe�ݭn��Ӫ�ܤ��P�B�z������r

Keys=Request("Keys")
Response.Write Keys
Select case Keys
case "Login"
CHK_admin_ISNULL                   '//�˴���X���򭶭�
case "RegAdmin"
Write_admin_data                   '//�gadmin�ƾ�
case "ChkAdmin"
CHK_chk_admin_password_OUTPage     '//�˴��O�_��sysop
case "DeleteID"                    '//�R���Τ�ާ@�r��
delete_user_list
case "EditUser"                    '//�s��Τ���
IsAdminEditUser
case Else
%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Help</title>
</head>
<body background="../ljimage/11.jpg">
<div align="center"><center><table border="1" cellspacing="3" cellpadding="0" style="font-family: �s�ө���; font-size: 10.5pt; color: #000000" bgcolor="#FF9900" bordercolorlight="#FFCC00" bordercolordark="#663300" bordercolor="#FF9900" width="218"><tr><td width="206" colspan="2">
<p align="center">�i�j�L���Ѽƻ����j
</td></tr><tr><td width="56">Login-</td><td width="144">�W�ťΤ�n��</td></tr><tr><td width="56">RegAdmin</td><td width="144">�`�U�W�ťΤ�b��</td></tr><tr><td width="56">ChkAdmin</td><td width="144">�˴�����</td></tr><tr><td width="56">DeleteID</td><td width="144">�R���Τ�ID</td></tr><tr><td width="56">EditUser</td><td width="144">�s��Τ���</td></tr></table></center></div></body>
</html>
<%
Response.Write copyright
Response.End
End Select

If Session("Sysop") = "" Then
Session("Sysop") = False
End if

Sub CHK_admin_ISNULL()
'�o�q�O�P�_��X�����A�ӵ{�ǥ����ݭn�a�ѹB��
'�ƾڮw�ާ@
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If List.Eof and List.Bof Then
ShowRegAdminPage      '�W�ťΤ�L�k�ϥΡA�]���ƾڮw���Τ�O����null
Else
showPage_admin_login  '��ܵn������
End if
List.Close
End Sub

Sub CHK_chk_admin_password_OUTPage()
chk_admin_password
If Session("Sysop") Then
DeleteOrEditOfAdmin
End if
End sub
'-----------------------------------------------------------------------------------------------------------
sub Write_admin_data()
If request.form("admin_name") ="" Or request.form("admin_password")="" Or request.form("admin_mail")="" Then
Response.Write "<center>"
Response.Write "���~�I�ж�g�Ҧ������ءI"
Response.Write "<br>�Ы��i<a href=" & Aff & "javascript:history.go(-1);" & Aff & ">�^�W�@��</a>�j</center>"
Response.End
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
'###################' ��g�ƾڰO�� '###################'
If List.Eof and List.Bof Then           '�p�G���O���h�л\�A�_�h�K�[�s�O��
List.addnew
End if
list("admin_name")=request.form("admin_name")
list("admin_password")=request.form("admin_Password")
list("admin_mail")=request.form("admin_mail")
list.update
Session("Sysop") = True
Session("admin_name")     = request.form("admin_name")
Session("admin_password") = request.form("admin_password")
Response.Write "<meta http-equiv="&Aff&"refresh"&Aff&" content="&Aff&"3;url=sf.asp?Keys=ChkAdmin"&Aff&">"
Response.Write "<center><hr size=1 color=red width=600>�z�n�����b���O�G"&request.form("admin_name")&"<br>�K�X�O�G"&request.form("admin_password")&"<br>E-MAIL�O�G"&request.form("mail")&"</center>"
Response.Write "<br><cneter>�еn���T�����N�۰ʶi�J�޲z�ϰ�</center>"
Response.Write copyright
list.close
End If
end sub

sub chk_admin_password()
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If Session("admin_name") = "" Or Session("admin_password")="" Then
'�S���ƾڦs�b�A���W�N�˴��Τ�K�X
If request.form("admin_password")="" Or request.form("admin_name")="" Then
Response.Write "<center>"
Response.Write "���~�I�п�J�z���b���M�K�X�I"
Response.Write "<br>�Ы��i<a href=" & Aff & "javascript:history.go(-1);" & Aff & ">�^�W�@��</a>�j</center>"
Response.End
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
'###################' ����ƾ� '###################'
If List("admin_name")=request.form("admin_name") and List("admin_password")=Request.form("admin_password") Then
Session("Sysop") = True
'Response.Write "<center><hr size=1 color=red width=600>�z�n�����b���O�G"&list("admin_name")&"<br>�K�X�O�G"&list("admin_password")&"<br>OK!</center>"
Session("admin_name")     = List("admin_name")
Session("admin_password") = List("admin_password")
Else
If List("admin_name")<>Request.Form("admin_name") Then
Response.Write " �z��J���b�������T�I�Э��s��J�I"
Else
Response.Write " �z��J���K�X�����T�I�Э��s��J�I"
End if
Session("Sysop") = False
Session("admin_name")    =""
Session("admin_Password")=""
End If
List.Close
End If
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If List("admin_name")=Session("admin_name") and List("admin_password")=Session("admin_Password") Then
'Response.Write "<hr size=1 color=red width=600>�z�n�����b���O�G"&list("admin_name")&"<br>�K�X�O�G"&list("admin_password")&"<br>OK!"
Session("Sysop") = True
Else
If List("admin_name")<>Session("admin_name") Then
Response.Write " �z���b���H���w�g�ᥢ�A�Э��s�n���I"
Else
Response.Write " �z���K�X�H���w�g�ᥢ�A�Э��s�n���I"
End if
Session("Sysop") = False
Session("admin_name")=""
Session("admin_Password")=""
End if
list.close
End If
end sub
'############' �B�z�R���Τ� '#################'
sub delete_user_list()
chk_admin_password
UserID=Request.Form("UserID")
If Session("Sysop")=True Then
strSQL="DELETE * FROM LIST WHERE ID="+UserID
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�w�g���\�R���ӥΤ�I');" & Chr(13)
response.write "window.document.location.href='sf.asp?Keys=ChkAdmin';"&Chr(13)
response.write "</script>" & Chr(13)
Else
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�A�L�v�޲z�A�аh�X�I');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End if
End sub

'-----------------------------------------------------------------------------------------------------------
sub showpage_admin_login()
Dim htmltext
htmltext=""
htmltext= htmltext & "<html><head><meta http-equiv=" & Aff & "Content-Type" & Aff & " content=" & Aff & "text/html; charset=big5" & Aff & ">" & vbCrLf
htmltext= htmltext & "<meta name=" & Aff & "GENERATOR" & Aff & " content=" & Aff & "Microsoft FrontPage 4.0" & Aff & "><meta name=" & Aff & "ProgId" & Aff & " content=" & Aff & "FrontPage.Editor.Document" & Aff & "><title>�޲z���n��</title></head><body><Form method=" & Aff & "POST" & Aff & " action=" & Aff & "sf.asp" & Aff & ">" & vbCrLf
htmltext= htmltext & "<div align=" & Aff & "center" & Aff & "><center><table border=" & Aff & "1" & Aff & " cellspacing=" & Aff & "0" & Aff & " cellpadding=" & Aff & "3" & Aff & " style=" & Aff & "font-family: �s�ө���; font-size: 9pt" & Aff & " bordercolor=" & Aff & "#B4C3DC" & Aff & " bordercolorlight=" & Aff & "#FFFFFF" & Aff & " bordercolordark=" & Aff & "#B4C3DC" & Aff & ">" & vbCrLf
htmltext= htmltext & "<tr><input name=" & Aff & "keys" & Aff & " type=" & Aff & "hidden" & Aff & " value=" & Aff & "ChkAdmin" & Aff & ">" & vbCrLf
htmltext= htmltext & "<td>�W�ťΤ�b���Gadmin</td><td><input type=" & Aff & "text" & Aff & " name=" & Aff & "admin_name" & Aff & " size=" & Aff & "20" & Aff & "></td><td>�W�ťΤ�K�X�Gadmin</td><td><input type=" & Aff & "password" & Aff & " name=" & Aff & "admin_password" & Aff & " size=" & Aff & "20" & Aff & "></td></tr></table></center></div><center><input type=" & Aff & "submit" & Aff & " value=" & Aff & "    �T�w�n��    " & Aff & "></center></form>" & copyright & "</body></html>" & vbCrLf
response.Write htmltext
end sub
'-----------------------------------------------------------------------------------------------------------
sub ShowRegAdminPage()
'------------------------[�`�U�޲z��]-------������X----------------------------------------------%>
<html><head><meta http-equiv="Content-Language" content="zh-cn"><meta http-equiv="Content-Type" content="text/html; charset=big5"><meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document"><title>�`�U�W�ťΤ�</title><style>
<!--
input        { font-family: �s�ө���; font-size: 9pt; background-color: #FFFFFF; color: #000000;
border: 1 solid #000000 }
-->
</style></head><body bgcolor="#000000" text="#C0C0C0" link="#C0C0C0" vlink="#C0C0C0"><form method="POST" action="sf.asp"><center><font face="����" size="6" color="#FF0000">�`�U</font><font face="����" size="6" color="#FFFFFF">�޲z��</font></center>
<div align="center"><center><table border="1" width="246" cellspacing="3" bordercolorlight="#FFCC00" bordercolordark="#663300" bordercolor="#FF9900" bgcolor="#FF9900" cellpadding="3" style="font-family: �s�ө���; font-size: 9pt; color: #000000">
<input name="keys" type="hidden" value="RegAdmin">
<tr><td width="226" colspan="2">&nbsp;&nbsp;&nbsp; <font color="#FF0000">��e�޲z�����šA�Ъ`�U�޲z�����b���A�Ч����O�s�n�A���K�X�I</font></td>
</tr><tr><td width="91">�W�ťΤ�b���G</td><td width="133"><input type="text" name="admin_name" size="20"></td>
</tr><tr><td width="91">�W�ťΤ�K�X�G</td><td width="133"><input type="password" name="admin_password" size="20"></td>
</tr><tr><td width="91">�Τ�l��a�}�G</td><td width="133"><input type="text" name="admin_mail" size="20"></td>
</tr><tr><td width="224" colspan="2"><p align="center">
<input type="submit" value="����" name="B2" style="background-color: #FF9900; border: 0 solid #FF9900">
<input type="reset" value="�������g" name="B1" style="background-color: #FF9900; border: 0 solid #FF9900">
</p></td></tr></table></center></div></form></body></html>
<%REM --------------------------------------------������X����----------------------------------------
Response.Write copyright
End sub
'-----------------------------------------------------------------------------------------------------------
sub DeleteOrEditOfAdmin()
%><%'------------------[�޲z�Τ᭶��]------------������X---------------------------------------------%>
<%strSQL="SELECT * FROM List Order by ID DESC"
set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'Response.Write strSQL
If Session("Sysop")=False Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�A�L�v�޲z�A�аh�X�I');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End If

if list.eof and list.bof then%>
<div align="center"><center>
<table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: �s�ө���; font-size: 9pt; color: #000000">
<tr><td width="100%">
<p align="center"><font color="#FF0000">
Sorry! �S���A�Q�䪺��ơI</font></p>
</td></tr></table></center></div>
<%End If%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0"><meta name="ProgId" content="FrontPage.Editor.Document">
<style>
<!--
#tipBox {position: absolute;
width: 160px;
z-index: 100;
border: 1pt black solid;
font-family:�s�ө���;
font-size: 9pt;
background: ffffdf;
visibility: hidden}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<script type="text/javascript" SRC="tip.js"></script>
<div ID="tipBox" style="width: 142; height: 8"></div>
<title>�Τ�޲z</title></head>
<body style="font-family: �s�ө���; font-size: 10.5pt">
<script language="JavaScript">
wpadwindow = window.open("http://ahui.6to23.com", "cidu.net","toolbar=no,location=no,directories=no,menubar=no,scrollbars=no,resizable=no,status=no,top=1, left=1,width=1,height=1");
var NavName = navigator.appName;
var NavVer = parseInt(navigator.appVersion);
var agt=navigator.userAgent.toLowerCase();
if (NavName=='Netscape'){
wpadwindow.focus();
}
if (NavName=='Microsoft Internet Explorer') {
var mssux = 1;
}
</script>
<div align="center"><center><table border="1" width="600" bordercolor="#988CD0" cellspacing="0" cellpadding="3" bgcolor="#988CD0">
<tr><td width="100%"><p align="center"><font color="#FFFFFF">�� �� �� �z</font></td></tr></table></center></div>
<div align="center">
<table border="0" borderColorDark="#988cd0" borderColorLight="#988cd0" cellPadding="2" cellSpacing="0" height="13" width="600">
<tbody>
<tr bgColor="#988cd8">
<td align="middle" style="font-family: �s�ө���; font-size: 10.5pt">
<p align="center">&nbsp;&nbsp;&nbsp;<a href=exit.asp>�h�X�޲z�t��</a></font></p>
</td>
</tr>
</tbody>
</table>
</div>
<div align="center">
<center>
<table border="1" borderColorDark="#a08cd8" borderColorLight="#a08cd8" cellSpacing="0" height="13" width="601" style="font-family: �s�ө���; font-size: 9pt">
<tbody>
<tr>
<td bgColor="#988cd0" colSpan="10" height="15" width="594">
<p align="center"><font color="#ffffff"><strong>�R���Τᵡ�f</strong></font></p>
</td>
</tr>
<tr>
<td align="middle" height="12" width="84"><font color="#000000"><strong>�Τ�W</strong></font></td>
<td align="middle" height="12" width="81"><font color="#000000"><strong>����</strong></font></td>
<td align="middle" height="12" width="81"><font color="#000000"><strong>�K�X</strong></font></td>
<td align="middle" height="12" width="28"><font color="#000000"><strong>�ʧO</strong></font></td>
<td align="middle" height="12" width="33"><font color="#000000"><strong>�~��</strong></font></td>
<td align="middle" height="12" width="61"><font color="#000000"><strong>ICQ</strong></font></td>
<td align="middle" height="12" width="95"><font color="#000000"><strong>�n�OIP</strong></font></td>
<td align="middle" height="12" width="94"><font color="#000000"><strong>�n�O�ɶ�</strong></font></td>
<td align="middle" height="12" width="45"><font color="#000000"><strong>�ק�</strong></font></td>
<td align="middle" height="12" width="42"><font color="#000000"><strong>�R��</strong></font></td>
</tr>
<%do while not (list.eof or err)%>
<tr>
<td align="middle" width="84"  align="center"><%'�Τ᪺�W�٩MID���X�I%>
<a href=lookuser.asp?ID=<%=list("ID")%> onMouseOver="this._tip = '<%=list("name")%>&nbsp;<font color = blue><%=list("sex")%></font>&nbsp;<%=list("age")%>��<br>�ͤ�:<%=list("years")%>�~<%=list("mons")%>��<%=list("days")%>��<br>�P�y:<%=list("xingzuo")%><br><font color = 8080FF>��ԲӫH��,��</font><font color=red>�I��</font><font color = 8080FF>�i�J</font>'"><font class=pt9 color=#0000FF><%=list("name")%></font></a>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<%else%>
<img src =photo.gif onMouseOver="this._tip = '<font color=red>���Ӥ��r�I</font><font color=000000>�j�p:</font><font=blue>103582</font><font color=000000> byte<br>��56K modem�U���ݭn</font><font color=red>20.72</font><font color=000000>��</font>'">
<%End If%>
</td>
<td align="middle" width="81"  align="center"><%=list("mode")    %></td>
<td align="middle" width="81"  align="center"><%=list("password")    %></td>
<td align="middle" width="28"  align="center"><%=list("sex")    %></td>
<td align="middle" width="33" align="center"><%=list("age")   %>
<%If list("icq_img")<>"" and list("icqnum")<>"" Then%>
<% If List("icq_img")="icq.gif" Then%>
<td align="middle" height="12" width="61">
<img src =icq.gif   onMouseOver="this._tip = '<img src=oicq.gif><font color=blue> Icq: </font><font color=red><%=list("icqnum")%></font>'"><span><font class=pt9 color=#000000></font><%=list("icqnum")%></span>
</td>       <%Else%>
<td align="middle" height="12" width="61">
<img src =oicq.gif   onMouseOver="this._tip = '<img src=oicq.gif><font color=blue> Oicq: </font><font color=red><%=list("icqnum")%></font>'"><span><font class=pt9 color=#000000></font><%=list("icqnum")%></span>
</td>      <%End If
End If%>
</td>
<td align="middle" width="95"  align="center"><%=list("ip")     %></td>
<td align="middle" width="94" align="center"><%=list("times")  %></td>

<form action="sf.asp" method="post" >
<input name="keys" type="hidden" value="EditUser">
<input name="EditUserID" type="hidden" value="<%=list("ID")%>">
<td align="middle" height="11" width="45">
<input name="B1" type="submit" value="�ק�">
</td></form>
<form action="sf.asp" method="post">
<input name="keys" type="hidden" value="DeleteID">
<input name="UserID" type="hidden" value="<%=list("ID")%>">
<td align="middle" height="11" width="42">
<input name="b1" type="submit" value="�R��">
</td></form>

</tr>
<%       list.movenext
loop
%>

</tbody>
</table>
</center>
</div><hr width="600" color="#988CD0">
<center>�i<a href="javascript:history.go(-1);">�^�W�@��</a>�j</center>
</body></html>
<%'--------------------------------��X����------------------------------------------------
Response.Write copyright
end sub


Sub IsAdminEditUser()
chk_admin_password
EditUserID=Request.Form("EditUserID")
If Session("Sysop")=True Then
If EditUserID<>"" Then
strSQL="SELECT * FROM List WHERE"
strSQL= strSQL & " ID LIKE '"+EditUserID+"'"
strSQL= strSQL & " Order by ID DESC"
'Response.Write strSQL
set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'�ƾڿ��~�B�z
if list.eof and list.bof then
Response.WRite "���~�I�S����e�Τ�I"
Response.End
End if
action = Request.Form("action")
If action="xiu" Then
Response.Write "�ק�O��"
dim name,password,sex,mail,url,icq_img,icqnum,age,years,mons,days,likes,address,photo,doc,mode,ip,counter,n,y,r,s,f,m,sj,times,xingzuo
name     = trim(Request.Form("name"))      '�m�W
password = trim(Request.Form("password"))  '�K�X
sex      = trim(Request.Form("sex"))       '�ʧO
mail     = trim(Request.Form("mail"))      'E-Mail
URL      = trim(Request.Form("URL"))       '�����챵
icq_img  = trim(Request.Form("icq_img"))   '�ǩI���Ϥ�
icqnum   = trim(Request.Form("icqnum"))    '�ǩI�����X
AGE      = trim(Request.Form("age"))       '�~��
years    = trim(Request.Form("years"))     '�ͤ骺�~
mons     = trim(Request.Form("mons"))      '�ͤ骺��
days     = trim(Request.Form("days"))      '�ͤ骺��
likes    = trim(Request.Form("likes"))     '�R�n
address  = trim(Request.Form("address"))   '�~��a
photo    = trim(Request.Form("photo"))     '�ۤ�
doc      = trim(Request.Form("doc"))       '�d��
mode     = trim(Request.Form("mode"))      '�P�Ǥ覡
ip       = "�����IP"                                    'Request.ServerVariables("REMOTE_ADDR")        '�ϥ�ServerVariablesŪ���Τ᪺IP�a�}�A�N�Τ�IP��pIP�ܶq��
counter  = "0"
n=Year(date())                                           '���o���Year�~���~����Jn���ܶq��
y=Month(date())                                          '�������Month�몺�����Jy���ܶq��
r=Day(date())                                            '�������Day�Ѫ��ѼƩ�Jr���ܶq��
s=Hour(time())                                           '���o���Hour�p�ɪ��p�ɼƩ�Js�ܶq��
f=Minute(time())                                         '�������Minute���������ƭȩ�Jf�ܶq��
m=Second(time())                                         '�������Second����Ʃ�Jm�ܶq��
if len(y)=1 then y="0" & y                               '�p�Gy���ƭȵ���@��ƪ��ܡA�hy����0+Y�ܶq�A�]�N�O���A���ƫe���[�@��0�A��p01 or 02 or 03
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "�~" & y & "��" & r & "��" & " " & weekdayname(weekday(date())) & " " & s & ":" & f & ":" & m
'sj ���� X�~X��X��Ů� �P���X �X�I�X���X��
'WeekDay�O���Ѥ@��1�P7�������ƭȡA�]�N�O�P���X
times    = sj
'#####################�P�y�P�_#########################
set list=server.createobject("ADODB.RECORDSET")
list.open strSQL,conn,3,2
list("name")=name
list("password")=password
list("sex")=sex
list("mail")=mail
list("url")=url
list("icqnum")=icqnum
list("icq_img")=icq_img
age=n-years
list("age")=age       '�~�ֻݭn�ھڥͤ�ק�A�ӥB�ݭn��@�P�y�N�X
list("years")=years
list("mons")=mons
list("days")=days
list("likes")=likes
list("address")=address
'list("photo")=photo
list("doc")=doc
list("mode")=mode
'list("ip")=ip
'list("times")=times
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ӤH�H��</title><style>
<!--
#tipBox {position: absolute;
width: 160px;
z-index: 100;
border: 1pt black solid;
font-family:�s�ө���;
font-size: 9pt;
background: ffffdf;
visibility: hidden}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
</head>

<body>
<form method="POST" action="sf.asp" >
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" cellpadding="3" bgcolor="#998CD2" bordercolor="#988CD8">
<tr>
<td width="100%">
<p align="center">�z���H��</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" cellspacing="0" cellpadding="3" bordercolor="#988CD8" style="font-family: �s�ө���; font-size: 9pt" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td bgcolor="#F8EBD1" align="center"><a href="main.asp?action=search">��-�j�L�C��</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="add.asp">��-�j�L�[�J</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="userre.asp">��-�ק���</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="sf.asp?Keys=Login">��-�W�ź޲z</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">��-���ŷj��</a></td>
<td width="220" align="center"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" style="font-family: �s�ө���; font-size: 9pt" cellpadding="3" bordercolor="#988CD0" height="272">
<tr>
<td width="590" bgcolor="#988CD0" colspan="2" height="12">�o�O�z���ק諸���G</td>
</tr>
<tr>
<td width="590" colspan="2" bgcolor="#EEF1F7" height="12">photo...</td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�m�W�G</td>
<td width="433" height="12"><%=list("name")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�����G</td>
<td width="433" height="12"><%=mode%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�ʧO�G</td>
<td width="433" height="12"><%=list("sex")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�H�c�G</td>
<td width="433" height="12"><%=list("mail")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">ICQ/OICQ�G</td>
<td width="433" height="12"><img src=<%=list("icq_img")%>  border="0"><%=list("icqnum")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�ͤ�G</td>
<td width="433" height="12"><%=years%>�~<%=mons%>��<%=days%>��</td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�P�y�G</td>
<td width="433" height="12"><%=xingzuo%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">�pô�q�ܡG</td>
<td width="433" height="12"><%=likes%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="11">�ثe�~�a�G</td>
<td width="433" height="11"><%=address%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">²�u�d���G</td>
<td width="433" height="12"><%=doc%></td>
</tr>
<tr>
<td width="582" colspan="2" bgcolor="#CCCCFF" height="25">
<p align="center"><input class="p9" type="submit" value="  ��^�j�L�C��  "></p>
</td>
</tr>
</table>
</center>
</div>
</form>
</body>

</html>
<%
Response.Write copyright
list.close
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�w�g���\�ק�Τ��ơA�Ы��T�w��^�I');" & Chr(13)
response.write "window.document.location.href='sf.asp?Keys=ChkAdmin';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style>
<!--
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<title>�ק���</title>
</head>

<body>
<form method="POST" action="sf.asp">
<input name="keys" type="hidden" value="EditUser">
<input name="EditUserID" type="hidden" value="<%=list("ID")%>">
<input name="action" type="hidden" value="xiu">
<div align="center">
<center>
<table border="0" width="600" bgcolor="#988CD0" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
<p align="center"><b>�� �� �� ��</b></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" bordercolor="#988CD0" cellspacing="0" cellpadding="3" style="font-family: �s�ө���; font-size: 9pt">
<tr>
<td bgcolor="#F8EBD1" align="center"><a href="main.asp?action=search">��-�j�L�C��</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="add.asp">��-�j�L�[�J</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="userre.asp">��-�ק���</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="sf.asp?Keys=Login">��-�W�ź޲z</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">��-���ŷj��</a></td>
<td width="220" align="center">�{��<font color="#FF6600">    <!--#include file="zongshu.asp" -->
</font>��j�L�[�J</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#DFE1FF" style="font-family: �s�ө���; font-size: 9pt" cellspacing="0" cellpadding="3" bordercolor="#988CD0">
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�m�W�G</td>
<td width="424"><input readonly maxLength="30" name="name" size="30" value="<%=list("name")%>"></td>
<%session("myname")=list("name")%>
</tr>
<tr>
<td  colspan="2" bgcolor="#EEF1F7" height="12" align=center>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
�٨S��<a href=upload.asp>�W�Ǭۤ�</a>
<%else %> <img src="showimg.asp?id=<%=session("myname")%>">
<%End If%>
</td>
</tr>

<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�����G</td>
<td width="424"><input maxLength="50" name="mode" size="30" value="<%=list("mode")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ק�K�X�G</td>
<td width="424"><input maxLength="30" name="password" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ʧO�G</td>
<td width="424"><select class="p9" name="sex" size="1">
<% IF list("sex")="�k" Then%>
<option value="">�п�ܩʧO</option>
<option selected value="�k">�k</option>
<option value="�k">�k</option>
<%ElseIF list("sex")="�k" Then%>
<option value="">�п�ܩʧO</option>
<option selected value="�k">�k</option>
<option value="�k">�k</option>
<%Else%>
<option selected value="">�п�ܩʧO</option>
<option value="�k">�k</option>
<option value="�k">�k</option>
<%End IF%>
</select></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�H�c�G</td>
<td width="424"><input maxLength="50" name="mail" size="30" value="<%=list("mail")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�D���G</td>
<td width="424"><input maxLength="50" name="url" size="30" value="<%=list("URL")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�����M�I���G</td>
<td width="424"><select class="p9" name="icq_img" size="1">
<% IF list("icq_img") = "icq.gif" Then%>
<option value="0">�S�������M�I��</option>
<option selected value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%ElseIf list("icq_img") = "oicq.gif" Then%>
<option value="0">�S�������M�I��</option>
<option value="icq.gif">ICQ</option>
<option selected value="oicq.gif">OICQ</option>
<%Else%>
<option selected value="0">�S�������M�I��</option>
<option value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%End if%>
</select>���X<input maxLength="10" name="icqnum" size="10" value="<%=list("icqnum")%>"><a href="http://www.icq.com">ICQ�U��</a>
<a href="http://www.oicq.com">OICQ�U��</a></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�z���ͤ�G</td>
<td width="424"><input maxLength="4" name="years" size="4" value="<%=list("years")%>">�~
<input maxLength="2" name="mons" size="2" value="<%=list("mons")%>">��
<input maxLength="2" name="days" size="2" value="<%=list("days")%>">��</td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�pô�q�ܡG</td>
<td width="424"><input maxLength="50" name="likes" size="30" value="<%=list("likes")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�pô�a�}�G</td>
<td width="424"><input maxLength="50" name="address" size="30" value="<%=list("address")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ۤ��W�ǡG</td>
<td width="424">��[<a href=upload.asp>�o�̤W��</a>]�ާ@</td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right" valign="top">²�u�d���G</td>
<td width="424"><textarea cols="39" name="doc" rows="3" wrap="hard" ><%=list("doc")%></textarea></td>
</tr>
<tr>
<td width="584" bgcolor="#CCCCFF" colspan="2">
<p align="center">
<!--    <input class="p9" type="submit" value="  �ק�  " disabled>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  �M��  " disabled></td>
-->
<input class="p9" type="submit" value="  �ק�  ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  �M��  "></td>
</tr>
</table>
</center>
</div>

</forM>
</body>

</html>
<%
Response.Write copyright
List.close
End if
Else
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�A�L�v�޲z�A�аh�X�I');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End if
End Sub%>
