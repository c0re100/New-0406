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
Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
Response.End
end if

username=info(0)
function amount()
dim rs
rs=conn.execute("Select sum(����q) As samount from �Ѳ�")
amount=rs("samount")
set rs=nothing
if isnull(amount) then amount=0
end function
%>
<!--#include file="jhb.asp"--><%
sql= "select * from �Ѳ�"
set rss=conn.execute(sql)
tai=rss("���A")

%>
<html>
<head>
<title>���Ʀ���Ѳ������</title>
<style type="text/css"><!--td {  font-family: �s�ө���; font-size: 9pt}body {  font-family: �s�ө���; font-size: 9pt}select {  font-family: �s�ө���; font-size: 9pt}A {text-decoration: none; font-family: "�s�ө���"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "�s�ө���"; font-size: 9pt} .big {  font-family: �s�ө���; font-size: 10.8pt}
--></style>
</head>

<body bgcolor="#FFFEF4">
<div align="center"><center>

<table border="0" width="100%" height="41" cellpadding="0" cellspacing="0">
<tr>
<td width="10%" height="16" bgcolor="#FFFAE1"><img border="0" src="logo.jpg" width="100" height="33"></td>
<td width="423" height="16" bgcolor="#FF6699">
<p align="center"><b><font color="#FBE651" face="����" size="3">���Ʀ���Ѳ������</font> </b>
</td>
<td width="10%" height="16" bgcolor="#FFFAE1"><img border="0" src="logo.jpg" width="100" height="33"></td>
</tr>
</table>
</center></div><div align="center"><center>

<table border="1" width="744" height="19" bordercolorlight="#FFFFFF" bordercolor="#000000"
cellspacing="0">
<tr>
<td width="744" height="19"><a href="jingji.asp">�Ѳ��g���H</a>---<a href="index.asp">�ѥ��污��s</a>---<a href="../jh.asp">��^
</a>
<%if info(2)>9  then%>
---<a href="kai.asp">�}�L</a>---<a href="feng.asp">����</a>---<a href="chufa.asp">�H���ƥ�</a>---<a href="jiu.asp">�ϥ�</a>---<a href="jiang.asp">���ѥ�����</a>
<%end if%>
</td>
</tr>
</table>
</center></div>

<table border="0" width="100%" bgcolor="#FFF0B2" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFAE1">
<td width="100%" height="20">����j�L��� -----�Ѳ��}���H�Ӧ����`�q<font
color="red"><%=amount()%></font>-----
�ثe�ѥ����A��<font
color="red"><%=tai%>�L</font><br>
�Ѳ��H��Ѫ��}�L������欰��ǡA���󬰺��A���󬰶^�A���^�H����q�M�H�Y���B�𬰼зǡC�����M�^�����d�򳣬�50%</td>
</tr>
</table>

<table border="0" width="100%" bgcolor="#FBE651" cellspacing="1" cellpadding="0">
<tr>
<td bgcolor="#006633" height="20" align="center" width="141"><font color="#FFFFFF">�Ѳ��N��</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">�Ѳ��W��</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">�}�L����</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">��e����</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">��/�^</font></td>
<td bgcolor="#006633" height="20" align="center" width="113"><font color="#FFFFFF">����q</font></td>
</tr>
<%
sql= "select * from �Ѳ�"
set rs=conn.execute(sql)
DO UNTIL RS.EOF  %>
<%if CSng(rs("�}�L����"))<CSng(rs("��e����")) then
color="#FF0000"
elseif CSng(rs("�}�L����"))>Csng(rs("��e����")) then
color="#00FF00"
else
color="#FFFF00"
end if

if (rs("��e����")-rs("�}�L����"))/rs("�}�L����")>=0.5 then
sstop=1
elseif (rs("��e����")-rs("�}�L����"))/rs("�}�L����")<=-0.5 then
sstop=2
else
sstop=0
end if

%>
<tr bgcolor="#000000">
<%if sstop=1 then%>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%>(����)</a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("���~")%></a></span></font></td>
<%elseif sstop=2 then%>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%>(�^��)</a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("���~")%></a></span></font></td>
<%else %>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%></a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("���~")%></a></span></font></td>
<%end if%>
<%
if rs("�y�q�Ѳ�")<=1000000 then
randomize time()
addstocknum=int(rnd*1000000+1000000)
end if
%>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("�}�L����"),2)%> </span></font></td>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("��e����"),2)%> </span></font></td>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("��e����")-rs("�}�L����"),2)%> </span></font></td>
<td height="20" width="113"><font color="<%=color%>"><span class="p"><%=rs("����q")%> </span></font></td>
</tr>
<%
rs.MoveNext
Loop
set rs=Nothing
%>
<%
sql="select * from �Ȥ� where �b��='"&username&"'"
set rs=conn.execute(sql)
if not(rs.eof or rs.bof) then
uname=rs("�b��")
%>
<table border="0" width="100%" bgcolor="#FFF0B2" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFAE1">
<td width="100%" height="20">��- �U���O�j�L<%=username%>�z�Ҿ֦����Ѳ��A<font color="red">���Ѳ����R��n�V���@�A�n���ե�1%���Ī��F�G�^</font></td>
</tr>
</table>
<%
set rs=nothing
%>
<table border="0" width="100%" bgcolor="#FBE651" cellspacing="1" cellpadding="0">
<tr>
<td bgcolor="#006633" height="20" align="center" width="141"><font color="#FFFFFF">�Ѳ��N��</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">�Ѳ��W��</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">�֦��ƶq</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">��������</font></td>
</tr>
<%
sql="select * from �j�� where �b��='"&uname&"' and ���Ѽ�<>0"
set rs_u=conn.execute(sql)
DO UNTIL RS_U.EOF

%>
<tr bgcolor="#000000">
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs_u("sid")%>"><font color=Violet><%=rs_U("sid")%></a></span></font> </td>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs_u("sid")%>"><font color=Violet><%=rs_U("���~")%></a></span></font> </td>
<td height="20" width="148"><font color=Violet><span class="p"><%=rs_u("���Ѽ�")%> </span></font></td>
<%if rs_u("��������")=0 then%>
<td height="20" width="148"><font color=Violet><span class="p"><%=formatnumber(rs_u("�R�J����"),3)%> </span></font></td>
<%else%>
<td height="20" width="148"><font color=Violet"><span class="p"><%=formatnumber(rs_u("��������"),3)%> </span></font></td>
<%end if%>
</tr>
<%
rs_u.MoveNext
Loop
end if
conn.close
set conn=nothing
%>

</table>
</body>
</html>