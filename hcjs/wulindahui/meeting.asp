<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")%>
<!--#include file="data.asp"-->
<html>
<head>
<title>�Z�L�j�|</title>
<style type="text/css">
td { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration: 
               underline }
</style>
</head>
<body bgcolor="#660000" text="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0">
<table border="0" height="24" width="623" cellspacing="0" cellpadding="0"> <tbody> 
<tr> <td height="38" width="623"> <table width="100%" border="0" cellspacing="0" cellpadding="0"
              bordercolorlight="#000000" bordercolordark="#FFFFFF"
              align="center"> <tr> <td width="91" height="26"><font size="2">&nbsp; 
<font
                    color="#FFFFFF"></font><font size="2"><font color="#ffffff"
                    size="2"><span class="zilong"><font color="#FFCC00"> <%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "��" & r & "��" %> </font></span></font></font></font></td><td height="26"> 
<div align="center"> </div><div align="center"> </div></td></tr> </table></td></tr> 
</tbody> </table><table width="623" border="0" cellspacing="0" cellpadding="0" height="298"> 
<tr> <td valign="top" width="616" height="305"> <div align="center"> <table width="100%" border="0" cellspacing="0" cellpadding="0" height="303"> 
<tr> <td align="center" height="47"> <img src="images/dtitle2.gif" width="200" height="45"></td><td height="47"> 
<div align="right"></div></td></tr> <tr valign="top" align="center"> <td colspan="2" height="249"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="141"> <tr> 
<td width="11">&nbsp;</td><td height="13" class="td01" WIDTH="599">&nbsp;</td><td width="10">&nbsp;</td></tr> 
<tr> <td class="td" WIDTH="11">&nbsp; </td><td valign="top" WIDTH="599"> <table width="97%" border="0" cellspacing="0" cellpadding="0" height="125"> 
<tr> <td align="center" height="35" colspan="2"><img src="images/gold.gif" width="120" height="40"></td><td height="35" align="center" colspan="2"><img src="images/silver.gif" width="120" height="40"></td><td height="35" align="center" colspan="2"><img src="images/copper.gif" width="120" height="40"></td></tr> 
<tr> <td width="15%" height="22" align="center">�����G <%
                        set rs=server.createobject("adodb.recordset")
			sql="select * from gold"
			rs.open sql,conn1,1
			%> <%=rs("�m�W")%> </td><td width="17%" height="22" align="center"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs.movenext%> </td><td width="16%" height="22" align="center">�����G <%
                         set rs1=server.createobject("adodb.recordset")
			 sql1="select * from silver"
			 rs1.open sql1,conn1,1
			 %> <%=rs1("�m�W")%> </td><td width="17%" height="22" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs1.movenext%> </td><td width="15%" height="22" align="center">�����G <%
                          set rs2=server.createobject("adodb.recordset")
			  sql2="select * from copper"
			  rs2.open sql2,conn1,1
			  %> <%=rs2("�m�W")%> </td><td width="20%" height="22" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs2.movenext%> </td></tr> <tr> <td width="15%" height="23" align="center">�]���G<%=rs("�m�W")%> 
</td><td height="23" align="center" width="17%"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs.movenext%> </td><td width="16%" height="23" align="center">�]���G<%=rs1("�m�W")%> 
</td><td width="17%" height="23" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs1.movenext%> </td><td width="15%" height="23" align="center">�]���G<%=rs2("�m�W")%> 
</td><td width="20%" height="23" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
<%rs2.movenext%> </td></tr> <tr> <td width="15%" height="22" align="center">����G<%=rs("�m�W")%></td><td height="22" align="center" width="17%"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td><td width="16%" height="22" align="center">����G<%=rs1("�m�W")%></td><td width="17%" height="22" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td><td width="15%" height="22" align="center">����G<%=rs2("�m�W")%></td><td width="20%" height="22" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td></tr> 
</table></td><td class="td" WIDTH="10">&nbsp;</td></tr> <tr> <td height="13" WIDTH="11">&nbsp;</td><td height="13" class="td01" WIDTH="599">&nbsp;</td><td WIDTH="10">&nbsp;</td></tr> 
</table><br> <table width="80%" border="0" cellspacing="0" cellpadding="0" height="49"> 
<tr> <td height="60"><font color="#FF0000">���򴣿�</font>�G �U��ӬD�]���j�L�A���ɺ]�ݪZ�\<font color="#FF0000">300</font>�H�U�A���Ⱥ]�ݪZ�\<font color="#FF0000">300��800</font>�A�����]�ݪZ�\<font color="#FF0000">800</font>�H�W�C�b���]�e�ФT��ӫ��A���]���\�i�H�N�W�r�d�b�o�̡A�b��ѫǳq���j�a�A�æ�����ӤH���D�Թﹳ�A�٦��@�w�����y�C���O�D�]�����\�|���@�w���l���A����O����,�j�L�@�իK���C^_^</td></tr> 
</table></td></tr> </table></div></td><td width="10" height="305">&nbsp;</td></tr> </table><div align="center"> 
</div>
</body>
</html>
<%
rs.close
set rs=nothing
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
conn1.close
set conn1=nothing
conn.close
set conn=nothing
%>