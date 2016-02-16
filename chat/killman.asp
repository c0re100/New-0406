<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
response.expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>菜單</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" leftmargin="0">
<table border="1" width="125" cellspacing="0" cellpadding="0" bgcolor="#336633" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="100%" height="28">
<div align="center"><font color="#ffffff" size="2">十大惡人</font></div>
</td>
</tr>
</table>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim sql
dim rs
dim filename
rs.open "select top 10 * from 用戶 where 道德<0 and 狀態<>'無' and 門派<>'官府' order by 道德",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>沒有可排行的對像 </p>"
else
%>
<table border="1" cellspacing="0" width="125" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
<tr bgcolor="#336633">
<td align="center" width="62"><font color="#FFFFFF">惡 人 谷</font></td>
<td align="center" width="36"><font color="#FFFFFF">道 德</font></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" bgcolor="#F7F7F7" width="62"><a href='seeyou.asp?username=<%=rs("姓名")%>'><%=rs("姓名")%></a></td>
<td align="center" bgcolor="#F7F7F7" width="36"><%=rs("道德")%> </td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>19 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>
 