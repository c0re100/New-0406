<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim page
page=request.querystring("page")
PageSize = 15
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "delete * from �H�R where �ɶ�<now()-5",conn,3,3
rs.open "Select ����,�ɶ�,����,���] From �H�R Order by �ɶ� DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>����R��</title>
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background="../images/8.jpg">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0" align="center">
<tbody>
<tr>
<td height="15" width="100%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>����R��</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
<table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table border="1" cellspacing="1" cellpadding="0" width="97%"
bordercolorlight="#EFEFEF">
<tr bgcolor="#FFFFFF">
<td width="20%">
<div align="center"><font color="#FF6600"> �� �` �� </font></div>
</td>
<td width="25%">
<div align="center"><font color="#FF6600"> �� �� </font></div>
</td>
<td width="19%">
<div align="center"><font color="#FF6600"> �� �H �� �� </font></div>
</td>
<td width="36%">
<div align="center"><font color="#FF6600"> �� �` �� �] </font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#FFCC99"  onMouseOut="this.bgColor='#FFCC99';"onMouseOver="this.bgColor='#DFEFFF';"> 
<td width="20%">
<div align="center"> <font color="#0000FF"><%=rs("����")%></font>
</div>
</td>
<td width="25%">
<div align="center"> <%=rs("�ɶ�")%> </div>
</td>
<td width="19%">
<div align="center"> <%=rs("����")%> </div>
</td>
<td width="36%">
<div align="center"> <%=rs("���]")%> </div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
sw=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[�@<font color="red"><b><%=pa%></b></font>��][<font
color="red"><b><%=sw%></b></font>�H���`]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="jl1.asp?page=<%=page-1%>">�W�@��</a>][��<%=page%>��][<a href="jl1.asp?page=<%=page+1%>">�U�@��</a>]</div>
</td>
</tr>
</table>
</table>
</div>
</body>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As ���H from �H�R")
musers=tmprs("���H")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>

