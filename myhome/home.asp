<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
%>
<!--#include file="homecheck.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<title><%=username%>�A�^�a�F�I</title>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"
bgcolor="#000000" text="#000000" background="../ljimage/11.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table border="0" height="17" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td height="26">
<div align="center"><font color="#FF0000">  <b><font size="2">�A�^��F�ۤv���a�A�o�̨S���M�I�A�A�i�H�n�n�𮧤@�U�F�I</font></b>
</font></div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
<table width="738" border="1" cellspacing="0" cellpadding="0" align="center" height="507" bordercolor="#000000">
<tr>
<td width="145" valign="top" height="351"><font size="2" color="#000000">��e�ɶ��G<%=hour(now)&":"&minute(now)&":"&second(now)%><br>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select �m�W,�ʧO,�t��,����,����,�Z�\,���O,��O,����,���s,�Ȩ�,���A,�D�w,����,�y�O from �Τ� where �m�W='"&info(0)&"'")%>
�w��<%=rs("�m�W")%>�^�a<br>
�ʧO:<%=rs("�ʧO")%><br>
�t��:<%=rs("�t��")%><br>
����:<%=rs("����")%><br>
����:<%=rs("����")%><br>
�Z�\:<%=rs("�Z�\")%><br>
���O:<%=rs("���O")%><br>
��O:<%=rs("��O")%><br>
����:<%=rs("����")%><br>
���s:<%=rs("���s")%><br>
�Ȩ�:<%=rs("�Ȩ�")%><br>
���A:<%=rs("���A")%><br>
�D�w:<%=rs("�D�w")%><br>
����:<%=rs("����")%><br>
�y�O:<%=rs("�y�O")%><br>
<%set rs=nothing
conn.close%>
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select house,bigarea,homeopen from userinfo where user='"&info(0)&"'")%>
�Ы�����:
<%if rs("house")="house1" then%>
�}��T��
<%end if%>
<%if rs("house")="house2" then%>
�@�륭��
<%end if%>
<%if rs("house")="house3" then%>
�@�뤽�J
<%end if%>
<%if rs("house")="house4" then%>
���ؤ��J
<%end if%>
<%if rs("house")="house5" then%>
���q�O��
<%end if%>
<%if rs("house")="house6" then%>
���اO��
<%end if%>
<br>
�Ҧb��:<%=rs("bigarea")%>��<br>
�a�����A�G
<%if rs("homeopen")="1" then%>
�j�����}��
<%else%>
�j���򳬵�</font><font color="#000000">
<%end if%>
</font></td>
<td valign="top" height="351">
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" width="100%"
height="307" align="center">
<tr>
<td align="center" colspan="2" valign="top" width="648" height="451">
<div align="left">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
height="21" align="center" bordercolorlight="#000000"
bordercolordark="#FFFFFF">
<tr align="center" valign="middle">
<td height="6" align="left"><font size="2" color="#000000">�A�^����
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select bigarea,house from userinfo where user='"&info(0)&"'")%>
<%=rs("bigarea")%>�Ϫ�
<%if rs("house")="house1" then%>
�}��p�T�θ̡A�o�O�A���a�C�A���g�٪��A���O�ܦn�I�n�h�V�O�@�I��M���u�@�]�n�`�N����I
<%end if%>
<%if rs("house")="house2" then%>
�p���и̡A�o�̬O�A���a�C�A�b�F�C����ߨ����[�A�b�B�̪ͭ����U�U�\�_�F�o�Ӥp�ΡI
<%end if%>
<%if rs("house")="house3" then%>
�@�뤽�J�̡A�o�̬O�A���a�C�A�O�F�C����̪��@��L�ȡA���@�w�����J~�۫H�g�L�A���V�O�A���[�N�i�H�h���j���Фl�̤F�I
<%end if%>
<%if rs("house")="house4" then%>
���ؤ��J�̡A�o�̬O�A���a�C�o�O�@�Һ}�G���Фl�A�|�B�z�S�X�a���ŷx�C�A�b�F�C����V�o����~~�~��V�O�a�I
<%end if%>
<%if rs("house")="house5" then%>
���q�O�ָ̡A�o�̬O�A���a�C�O�֥|�P�����y�H~���H�̬��A�ݯ�����I�A�O�F�C����̪���O���A�֤]�����H�K�S�A�I
<%end if%>
<%if rs("house")="house6" then%>
���اO�ָ̡A�o�̬O�z������C�ޮa�a�ۤj�����H</font>
<div id="Layer2"
style="position:absolute; width:200px; height:115px; z-index:2; left: 448px; top: 159px">
<font size="2" color="#000000"><img src="christmas1.jpg"
width="279" height="181"></font> </div>
<font size="2" color="#000000">����z�ﱵ�i�F���U�A���z�e�W�F�W�Q�M���I�z�ݵ۸g�L�V�O�ӱo�쪺�@���A�ߤ��P�n�U�d�I�z�O�F�C����̪��F�C����H���I</font>
<font color="#000000">
<%end if%>
</font></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
<tr align="center" valign="middle">
<td height="5"></td>
</tr>
</table>
<div id="Layer1"
style="position:absolute; width:119px; height:91px; z-index:1; left: 241px; top: 199px">
<font color="#000000"><img src="a1.gif" width="118" height="88">
</font></div>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</body>
</html>
