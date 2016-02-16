<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>排行掌門TOP10</title>
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

const MaxPerPage=10
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
sql="select 姓名,門派,身份,武功,攻擊,防御,等級,體力 from 用戶 where 身份='掌門' and 門派<>'官府' order by 武功+攻擊+防御 desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
 
response.write "<p align='center'><font color=ffffff>沒有可排行的對像</font></p>"
else
%>
  <table border="1" cellspacing="1" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
    <tr bgcolor="#990000"> 
      <td align="center" height="17" width="12%"><font size="2" color="#FFFFFF">姓 
        名</font></td>
      <td align="center" height="17" width="12%"><font size="2" color="#FFFFFF">門 
        派</font></td>
      <td align="center" height="17" width="12%"><font size="2" color="#FFFFFF">身 
        份</font></td>
      <td align="center" height="17" width="12%"><font size="2" color="#FFFFFF">武功</font></td>
      <td align="center" height="17" width="9%"><font size="2" color="#FFFFFF">攻擊</font></td>
      <td align="center" height="17" width="10%"><font size="2" color="#FFFFFF">防御</font></td>
      <td align="center" height="17" width="10%"><font size="2" color="#FFFFFF">等級</font></td>
      <td align="center" height="17" width="14%"><font color="#FFFFFF">合計</font></td>
      <td align="center" height="17" width="19%"><font size="2" color="#FFFFFF">體 
        力</font></td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td align="center" width="12%"><font color="#FFFFFF"><%=rs("姓名")%></font> 
      </td>
      <td align="center" width="12%"><font color="#CCFFCC"><%=rs("門派")%></font> 
      </td>
      <td align="center" width="12%"><font color="#CCFFCC"><%=rs("身份")%></font></td>
      <td align="center" width="12%"><font color="#CCFFCC"><%=rs("武功")%></font> 
      </td>
      <td align="center" width="9%"><font color="#CCFFCC"><%=rs("攻擊")%></font></td>
      <td align="center" width="10%"><font color="#CCFFCC"><%=rs("防御")%></font></td>
      <td align="center" width="10%"><font color="#CCFFCC"><%=rs("等級")%></font></td>
      <td align="center" width="14%"><font color="#FFDDCC"><%=rs("武功")+rs("攻擊")+rs("防御")%></font></td>
      <td align="center" width="19%"><font color="#CCFFCC"><%=rs("體力")%></font></td>
    </tr>
    <%
rs.movenext
filename=filename+1
if filename>9 then Exit Do
loop
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 
%></table>
  <br>
</div>
</body>
</html>
 