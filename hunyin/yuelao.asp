<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
myname=request.cookies("Jname")
myname=info(0)
rs.open "delete * from ��� where �ɶ�<now()-5",conn,3,3
rs.open "SELECT �m�W1,�m�W2,����,�ɶ� FROM ��� order by �ɶ� DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>���</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style>td           { font-size: 14; color: #000000 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" bgcolor="#660000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<a href="zhenfen.asp">�ڭn�n�O</a><font color="#0000FF"><br>

<a href="#" onClick="window.open('yuanou.asp','lihun','resizable=no,width=460,height=120,scrollbars=no,Status=no,toolbar=no,left=300,top=300,menubar=no,fullscreen=no')">����谸</a>
</font> <br>
<font color="#0000FF"><a href="Stunt.asp" title="���B�n�l���ͩR100"> �X��S��</a> </font>
<table width="550" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
<tr bgcolor="#f7f7f7">
    <td align="left"><font size="2">[�@<font color="#990000"><b><%=rs.pagecount%></b></font>��][<font
color="#990000"><b><%=musers()%></b></font>�H�D�B] [<a
href="yuelao.asp?page=<%=page-1%>"><font color="#990000">�W�@��</font></a>][��<%=page%>��][<a
href="yuelao.asp?page=<%=page+1%>"><font color="#990000">�U�@��</font></a>]</font></td>
</tr>
<tr>
<td style="font-size:21;color:#FF1493" align="center"><font size="2"><b><font size="+2">�D�B�H��</font></b></font></td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding="2" width="100%"
bgcolor="#000000" bordercolorlight="#EFEFEF">
<tr bgcolor="#DFEDFD">
<td><font size="2">�D�B��</font></td>
<td><font size="2">�D�B�ﹳ</font></td>
<td><font size="2">�u�����</font></td>
<td><font size="2">�ɶ�</font></td>
<td><font size="2">�^�_</font></td>
</tr>
<%
count=0
do while not rs.eof and count<rs.PageSize
%>
<tr bgcolor="#f7f7f7">
<td><font size="2"><%=rs("�m�W1")%></font></td>
<td><font size="2"><%=rs("�m�W2")%></font></td>
<td><font size="2"><%=rs("����")%></font></td>
<td><font size="2"><%=rs("�ɶ�")%> </font>
<td><font size="2">
<%if rs("�m�W2")=myname then %>
<a
href="jiefen.asp?name=<%=rs("�m�W1")%>">�P�N</a>
<%end if%>
</font></td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</td>
</tr>
</table>
</body>

</html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �m�W1 from ���")
musers=tmprs("�m�W1")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%> 