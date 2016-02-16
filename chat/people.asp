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
sql="select * from 文本 "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<html>
<head>
<title>值班人員</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="0" bgcolor="#660000">
<table width="135" border="0" height="42%" align="center">
  <tr align="center">
    <td height="370"> 
      <p align="left"><font size="2"><font color=#CCFFCC size=2>官府人員值班時間表:<br>
        <br>
        <%=rs("PK表")%></font></font></p>
      <div align="right"><br>
      </div>
    </td>
</tr>
</table>
</body>
</html> <%rs.movenext
loop
	set rs=nothing	
	conn.close
	set conn=nothing %>