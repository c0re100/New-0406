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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ʧO,�Ȩ� from �Τ� where id="&info(9),conn
sex=rs("�ʧO")
yin=rs("�Ȩ�")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
p{font-size:9pt; color:#000000}
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
<p align="center"><b><font style="font-size: 9pt" color="#ffee00">�w��<%=my%>�^��ۤv���p��</font></b><br><br>
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
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="�k" then%><img src="image/logonan.jpg"><%end if%><%if sex="�k" then%><img src="image/logonv.jpg"><%end if%></td>                 
  <td height="1" width="380" bgcolor="#C4DEFF" rowspan="10"> 
    <%if lx="���اO��" then%>
    <img src="image/35028_2.jpg" width="436" height="279"> 
    <%end if%>
  </td>                 
<td height="13" width="159" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">���w�̦�</font><font color="#ff0000"><%if sex="�k" then%><%=rs("�Ȩ�")%><%end if%><%if sex="�k" then%><%=rs("��Q�Ȩ�")%><%end if%></font>��</p>
</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="�k" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="�k" then%><form method="post" action="putmoney.asp"><%end if%>
�ڭn�s<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>��</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="�T�w"><INPUT TYPE="reset" VALUE="���g"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="�k" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="�k" then%><form method="post" action="putmoney.asp"><%end if%>
�ڭn��<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>��</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="�T�w"><INPUT TYPE="reset" VALUE="���g"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">�{�b���@�D���n�ڿ�</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">�٬O��b�a�̤���w</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">���ǡC</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">�G�A�^�ӤF</font></p>
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">��e��m�G���w</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../jh.asp"><img src="image/Bd_wd.GIF" width="40" height="20" border="0"></a></p>
</td>
<td height="164" width="157" bgcolor="#C4DEFF">
<font color="#ff0000">�z���W���Ȩ�G<%=yin%></font>
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

