<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp"  -->
<%
Dim USERID
Dim strSQL
Dim List,action
Dim name,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei,repassword
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,xingzuo
USERID=Request.Querystring("USERID")
IF Session("username")="" or Session("password")="" or USERID="" Then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�Τ�H���ᥢ�A�Э��s�n���I');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if

strSQL="SELECT * FROM List WHERE"
strSQL = strSQL & " name LIKE '%" & Session("username") & "%'"
strSQL = strSQL & " and password LIKE '%" & Session("password") & "%'"
strSQL = strSQL & " and ID LIKE '%" & USERID & "%'"
strSQL =strSQL & " Order by ID DESC"

set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'�ƾڿ��~�B�z
if list.eof and list.bof then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
Response.WRite "���~�I�S����e�Τ�I�Э��s�n���I"
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if

IF  Session("Username")<>list("name") and Session("password")<>list("password") and USERID<>list("id") Then             '�p�G�Τ�W�ٵ���NULL�h�ҩ��S���q�L����
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�S���n���A�Э��s�n���I');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if
session("mynameup")=list("name")
Response.Write "OK!�q�L�T�{�F..."                        '�T�{�Τ�w�g�q�L�T�{�I
Response.Write Session("Username")

action = Request.Querystring("action")
If action="xiu" Then
Response.Write "�ק�O��"
name     = trim(Request.Form("name"))      '�m�W
password = trim(Request.Form("password"))  '�K�X
repassword = trim(Request.Form("repassword"))  '�K�Xre
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
gsm      = trim(Request.Form("gsm"))       '���
pc  	 = trim(Request.Form("pc"))   	   '�l�F�s�X
danwei   = trim(Request.Form("danwei"))    '���
address  = trim(Request.Form("address"))   '�~��a
photo    = trim(Request.Form("photo"))     '�ۤ�
doc      = trim(Request.Form("doc"))       '�d��
mode     = trim(Request.Form("mode"))      '�P�Ǥ覡
ip       = Request.ServerVariables("REMOTE_ADDR")
if repassword<>password  then _
response.write "�K�X���~�C<a href = javascript:history.back()>��^����</a>":response.end

'�ϥ�ServerVariablesŪ���Τ᪺IP�a�}�A�N�Τ�IP��pIP�ܶq��
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
list("gsm")=gsm
list("pc")=pc
list("danwei")=danwei
list("address")=address
'list("photo")=photo
list("doc")=doc
list("mode")=mode
list("ip")=ip
list("times")=times
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ӤH�H��--<%=List("name")%></title><style>
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

<body background="../ljimage/11.jpg">
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
<td width="66" align="center"><a href=main.asp?action=search>��-�j�L�C��</a></td>
<td width="66" align="center"><a href="add.asp">��-�j�L�[�J</a></td>
<td width="66" align="center"><a href="userre.asp">��-�ק���</a></td>
<td width="66" align="center"><a href="sf.asp?Keys=Login">��-�W�ź޲z</a></td>
<td width="66" align="center"><a href="admsearch.asp?keys=admsearch">��-���ŷj��</a></td>
<td width="220" align="center"><font color="#996600">�{��</font><font color="#FF6600">
<!--#include file="zongshu.asp" --> </font><font color="#996600">��j�L�[�J</font></td>
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
      <!-- <tr>
<td width="590" colspan="2" bgcolor="#EEF1F7" height="12">photo...</td>
</tr> -->
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
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">�ͤ�G</td>
        <td width="433" height="12"><%=years%>�~<%=mons%>��<%=days%>��</td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">�pô�q�ܡG</td>
        <td width="433" height="12"><%=likes%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="11">�l�s/�a�}�G</td>
        <td width="433" height="11"><%=address%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">²�u�d���G</td>
        <td width="433" height="12"><%=doc%></td>
      </tr>
      <tr> 
        <td width="582" colspan="2" bgcolor="#CCCCFF" height="25"> 
          <p align="center">
            <input class="p9" type="submit" value="  ��^��ͦC��  ">
          </p>
        </td>
      </tr>
    </table>
</center>
</div>
</body>

</html>
<%
list.close
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�w�g���\�ק�Τ��ơA�Ы��T�w��^�I');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
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
<title>�ק���--<%=List("name")%></title>
</head>

<body>
<form method="POST" action="moduser.asp?USERID=<%=USERID%>&action=xiu">
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
<td  align="center"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='�ۤ��W��'>��-�ۤ��W��</a></td>
</tr>
</table>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<a href style='cursor:hand;'onClick="window.open('upload.asp','_top','scrollbars=yes,resizable=yes,width=600,height=400')" title='�W�Ǭۤ�'>�٨S���W�Ǭۤ�</a>
<%else %> <font size=2><a href style='cursor:hand;'onClick="javascript:location.reload();" title='��s'>�z��谷�</a><img src="showimg.asp?id=<%=session("Username")%>"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='�ۤ����s�W��'>�ڭn�ק�</a>
<%End If%>

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
<td width="160" bgcolor="#CCCCFF" align="right">�����G</td>
<td width="424"><input maxLength="30" name="mode" size="30" value="<%=list("mode")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ק�K�X�G</td>
<td width="424"><input maxLength="30" name="password" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�T�{�ק�K�X�G</td>
<td width="424"><input maxLength="30" name="repassword" size="10" type="password" value="<%=list("Password")%>"></td>
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
<a href="http://www.tencent.com">OICQ�U��</a></td>
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
<td width="160" bgcolor="#CCCCFF" align="right">����ΩI���G</td>
<td width="424"><input maxLength="50" name="gsm" size="30" value="<%=list("gsm")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�l�F�s�X�G</td>
<td width="424"><input maxLength="6" name="pc" size="30" value="<%=list("pc")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">���W�١G</td>
<td width="424"><input maxLength="50" name="danwei" size="30" value="<%=list("danwei")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ԲӦa�}�G</td>
<td width="424"><input maxLength="50" name="address" size="30" value="<%=list("address")%>"></td>
</tr>
<!--    <tr>
<td width="160" bgcolor="#CCCCFF" align="right">�ۤ��W�ǡG</td>
<td width="424"><input readonly name="photo" value="<%=list("photo")%>" type="file" disabled >&nbsp;
��JDelete�h�R���z���ۤ�</td>
</tr> -->
<tr>
<td width="160" bgcolor="#CCCCFF" align="right" valign="top">²�u�d���G</td>
<td width="424"><textarea cols="39" name="doc" rows="3" wrap="hard" ><%=list("doc")%></textarea></td>
</tr>
<tr>
<td width="584" bgcolor="#CCCCFF" colspan="2">
<p align="center">
<input class="p9" type="submit" value="  �ק�  ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  �M��  "></td>
</tr>
</table>
</center>
</div>

</forM>

</body>

</html>
