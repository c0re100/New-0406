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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT ���~�W,�ƶq FROM ���~ where  �֦���="&info(9)
Set Rs=conn.Execute(sql)
name=info(0)
%>
<html>
<head>
<title>�� �� �p ��</title>
<style>
p{font-size:9pt; color:#ffee00}
td,select,input{font-size:9pt; color:#000000; height:9pt}
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
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=700 height="282">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="11"><%if sex="�k" then%><img src="image/logonan.jpg"><%end if%><%if sex="�k" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="11"><%if lx="�@��Ы�" then%><img src="image/ccs-01.jpg"><%end if%><%if lx="���Ť��J" then%><img src="image/ccs-02.jpg"><%end if%><%if lx="���v��" then%><img src="image/ccs-03.jpg"><%end if%><%if lx="���اO��" then%><img src="image/ccs-04.jpg"><%end if%> </td>                 
<td height="1" width="200" bgcolor="#C4deff"><center>���~�W</center></td>
<td height="1" width="159" bgcolor="#C4deff"><center>�ƶq</center></td>

</tr>
<% 
do while not rs.eof and not rs.bof 
%>
<tr>
<td height="1" width="200" bgcolor="#C4DEFF"><center><%=rs("���~�W")%></center></td>
<td height="1" width="159" bgcolor="#C4DEFF"><center><%=rs("�ƶq")%></center></td>
           
</tr>
<% 
rs.movenext 
loop 
%>
</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=700 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
    <p align="center"><a href="../hcjs/es/wupin.asp" target="_blank"><font color="#000000">�ڪ��F��</font></a></p> 
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">���e��m�G�x�ë�</font></p>
</td>                 
<td height="164" width="78" bgcolor="#C4DEFF">
  <p align="center"><a href="../jh.asp"><img src="image/Bd_wd.GIF" width="40" height="20" border="0" target="_blank"></a></p>
</td>
<td height="164" width="78" bgcolor="#C4DEFF">
  <p align="center"><a href="../hcjs/es/wupin.asp" width="100" height="400" target="_blank"><font color="#000000">�H���]�q</font></a></p>
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