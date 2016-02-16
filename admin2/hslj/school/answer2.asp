<%
Response.Expires=0
if session("Ba_jxqy_username")="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>問題回答</title>
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
response.write "第"&I&"題"
if ann=an then
response.write "正確<br>"
r=r+1
else
response.write "錯！答案是"&ann&"<br>"
w=w+1
end if
next
jingli=10
zizhi=3
set conn=server.createobject("adodb.connection")
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用戶 set 銀兩=銀兩+"&jingli*r&",精力=精力+"&jingli*r&",資質=資質+"&zizhi*r&" where 姓名='"&Session("myname")&"'"
conn.Execute(sql)
response.write "<P>您一共：<br>答對了"&r&"題<br>答錯了"&w&"題<br>增加了"&jingli*r&"兩銀子，增加了"&jingli*r&"點精力，"&zizhi*r&"點資質！"
%></td>
</tr>
</body>





