<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
ipdizi=request.form("ipdizi")
if ipdizi="" then%>
<html>
<head>
<title>�F�C����</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor=#FFFFFF background="0064.jpg">
<center>
  <font color="#000000"><b><font size="+6" color="#FF0000"> 
  <%ip=Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'�d�̫�ip���a��
sql="select * FROM �a�} WHERE ip1<="& num &" and ip2>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="�Ȭw"
city="����"
else
country=rs("��a")
city=rs("����")
end if
if country="" then country="����"
if city="" then city="����"
last="�a��:"& country &" ����:"& city
set rs=nothing	
set conn=nothing
%>
�F�C����ip�d��{��</font></b></font><br>
<br>
�P�°l���@�̴��Ѫ�ip�ƾڮw�{��(11.1��e)<br>
�`�G�ӵ{�ǶȦ��J�ꤺip�a�}�I<br>
�z��ip�a�}���G<%=ip%> <%=last%>
<form action=ipseek.asp method=post>
�п�Jip�a�}�æ^��:
<input  size=20 name=ipdizi><input type=submit value='�T�{'>
</form>
</center>
</body>
</html>
<%else
ip=ipdizi
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'�d�̫�ip���a��
sql="select top 1 ��a,���� FROM �a�} WHERE ip1<="& num &" and ip2>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="�Ȭw"
city="����"
else
country=rs("��a")
city=rs("����")
end if
if country="" then country="����"
if city="" then city="����"
last="�a��:"& country &" ����:"& city
set rs=nothing	
set conn=nothing%>
<html>
<head>
<title>�F�C����</title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
�z�Ҭd�䪺ip�a�}���G<%=ip%><br>�D�w���G<%=last%>
<br>
<br>
<a href="ipseek.asp">��^</a>
</html>
<%end if%> 