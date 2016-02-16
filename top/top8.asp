<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>門派基金排行</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
const MaxPerPage=30
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
<body text="#000000" bgcolor=000000 topmargin="0">
<div align="center">
<%
dim sql
dim rs
dim filename
sql="select 門派,掌門,弟子數,幫派基金 from 門派 where 門派<>'官府' and 門派<>'遊俠' order by 幫派基金 desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then

response.write "<p align='center'>沒有可排行的對像 </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr bgcolor="#990000">
<td align="center" height="16" width="13%"><font color="#FFFFFF">門派名</font></td>
<td align="center" height="16" width="7%"><font color="#FFFFFF">掌門</font></td>
<td align="center" height="16" width="9%"><font color="#FFFFFF">弟子數</font></td>
<td align="center" height="16" width="9%"><font color="#FFFFFF">門派基金數</font></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" width="13%"><font color="#FFFFFF"><%=rs("門派")%></font></td>
<td align="center" width="7%"><font color="#CCFFCC"><%=rs("掌門")%></font></td>
<td align="center" width="9%"><font color="#CCFFCC"><%=rs("弟子數")%></font></td>
<td align="center" width="9%"><font color="#FFDDCC"><%=rs("幫派基金")%></font> </td>
</tr>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing

%></table>
</div>
<div align="center"></div>
</body>
</html>

