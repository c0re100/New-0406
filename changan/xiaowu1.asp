<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=info(0)
sex=info(1)
%>
<!--#include file="data1.asp"--><%
sql="select * from �Ы� where ��D='" & my & "' or ��Q='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z�٨S���R�Фl�O�I');location.href = 'fangwu.asp';}</script>"
elseif rs("���A")<>"���`" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z���ЫάO���O���I���p�r�I');location.href = '../jh.asp';}</script>"
else
lx=rs("����")
%>
<html>
<head>
<title>�� �� �p ��</title>
<style>
p{font-size:9pt; color:#ffee00}
td,select,input{font-size:9pt; color:#000000; height:14pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body bgcolor=#990099>
<center>
<p align="center"><b><font style="font-size: 9pt">�w��<%=my%>�^��ۤv���p��</font></b><br><br>
<TABLE width='95%' ALIGN=center CELLSPACING=2 BORDER=2 CELLPADDING=5 BGCOLOR='#90c088'><tr><td>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>
<tr bgcolor="#C4DEFF">
<td align="center" width=10%><a href="xiaowu.asp"><font color="#000000"><%if lx="�@��Ы�" or lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>���U</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu1.asp"><font color="#000000"><%if lx="�@��Ы�" or lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>�׫�</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu2.asp"><font color="#000000"><%if lx="�@��Ы�" or lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>�x�ë�</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu3.asp"><font color="#000000"><%if lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>�\�U</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu4.asp"><font color="#000000"><%if lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>�åͶ�</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu5.asp"><font color="#000000"><%if lx="���Ť��J" or lx="���v��" or lx="���اO��" then%>�p���w</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu6.asp"><font color="#000000"><%if lx="���v��" or lx="���اO��" then%>���</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu7.asp"><font color="#000000"><%if lx="���اO��" then%>�C�a��</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu8.asp"><font color="#000000"><%if lx="���اO��" then%>������</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu9.asp"><font color="#000000"><%if lx="���v��" or lx="���اO��" then%>�ѩ�</font></a><%end if%></td>
</tr></table></TABLE>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="281">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="11"><%if sex="�k" then%><img src="image/logonan.jpg"><%end if%><%if sex="�k" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="11"><%if lx="�@��Ы�" then%><img src="image/ws-01.jpg"><%end if%><%if lx="���Ť��J" then%><img src="image/ws-02.jpg"><%end if%><%if lx="���v��" then%><img src="image/ws-03.jpg"><%end if%><%if lx="���اO��" then%><img src="image/ws-04.jpg"><%end if%> </td>                 
<td height="1" width="159" bgcolor="#C4DEFF">
<form method=POST action='pub1.asp'>
�A�Q�n�𮧡G</td>                 
</tr>
<tr><td height="1" width="159" bgcolor="#C4DEFF">
<select name=date size=1>
<option value="0">�s��
<option value="1">�@��
<option value="2">���
<option value="3">�T��
</select>
<select name=time size=1> 
<option value="0">0�p��
<option value="1">1�p��
<option value="2">2�p��
<option value="3">3�p��
<option value="4">4�p��
<option value="5">5�p��
<option value="6">6�p��
<option value="7">7�p��
<option value="8">8�p��
<option value="9">9�p��
<option value="10">10�p��
<option value="11">11�p��
<option value="12">12�p��
<option value="13">13�p��
<option value="14">14�p��
<option value="15">15�p��
<option value="16">16�p��
<option value="17">17�p��
<option value="18">18�p��
<option value="19">19�p��
<option value="20">20�p��
<option value="21">21�p��
<option value="22">22�p��
<option value="23">23�p��
</select>
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
<INPUT TYPE="submit" VALUE="�T�w"><INPUT TYPE="reset" VALUE="���g">
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">�b�׫Ǹ̥𮧨C�p��</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">�[���O1000��O1000</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">���ݭn�Ȩ�C</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF"> </td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
    <p align="center">&nbsp;</p>
  </td>                 
</tr></form></table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
<p align="center"><a href="../myhome/sheepboy/feedsheep.asp" target="_blank"><font color="#000000">�ӮƫĤl</font></a></p> 
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">��e��m�G�׫�</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../jh.asp"><img src="image/Bd_wd.GIF" width="40" height="20" border="0"></a></p>
</td>                 
</tr>
</table>
<br><br><br><br><br><!--#include file="copy.inc"-->
</body></html>
<%end if
set rs=nothing
conn.close
set conn=nothing
%>