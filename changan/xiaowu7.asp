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
td,select,input{font-size:9pt; color:#000000; height:9pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #000000; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #000000;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
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
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="�k" then%><img src="image/logonan.jpg"><%end if%><%if sex="�k" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="10"><%if lx="���اO��" then%><img src="image/yyc-01.jpg"><%end if%> </td>                 
<td height="13" width="159" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000"><a href="youyong.asp">�C�a</a></font></p>
</td>                 
</tr>
<tr>
<td height="10" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="15" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="15" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="19" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="24" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="37" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="50" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="83" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="242" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">�G�A�^�ӤF</font></p>
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">��e��m�G�C�a��</font></p>
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