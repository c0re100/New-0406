<%
Response.Expires=0
if session("Ba_jxqy_username")="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>���D�^��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style4.css" rel=stylesheet>
</head>

<body oncontextmenu=self.event.returnValue=false  body bgcolor="<%=Application("Ba_jxqy_backgroundcolor")%>
<table width=" border="0" cellspacing="0" cellpadding="0" align="center" 595">
<table>
<tr>
<td><%dim r,w
r=0
w=0
for i=1 to 10
ann=request.form("ann"&I&"")
an=request.form("an"&I&"")
response.write "��"&I&"�D"
if ann=an then
response.write "���T<br>"
r=r+1
else
response.write "���I���׬O"&ann&"<br>"
w=w+1
end if
next
jingli=10
zizhi=3
set conn=server.createobject("adodb.connection")
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="update �Τ� set �Ȩ�=�Ȩ�+"&jingli*r&",��O=��O+"&jingli*r&",���=���+"&zizhi*r&" where �m�W='"&Session("myname")&"'"
conn.Execute(sql)
response.write "<P>�z�@�@�G<br>����F"&r&"�D<br>�����F"&w&"�D<br>�W�[�F"&jingli*r&"��Ȥl�A�W�[�F"&jingli*r&"�I��O�A"&zizhi*r&"�I���I"
%></td>
</tr>
</body>





