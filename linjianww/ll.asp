<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
'if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
'if info(2)<10 then Response.Redirect "../error.asp?id=900"
larenseek=Request.Form("larenseek")
%>
<html>
<head>
<title>會員查詢程序</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="../images/8.jpg">
<p align="center"> <font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font>
<br>
感謝你朋友這些人是你拉到我們江湖的！<br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">江湖記錄</font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="470" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 單挑 "
Set Rs=conn.Execute(sql)
%>
          <td width="86" height="10"> 
            <div align="center"><font color="#FFFFFF">身份</font></div>
          </td>
          <td width="75" height="10"> 
            <div align="center"><font color="#FFFFFF">最後登陸時間</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="86" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("姓名1")%></font></div>
          </td>
          <td width="75" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("姓名2")%></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
</td>
</tr>
</table>
<div align="center"><font color="#000000">拉人總數:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font>
</div>
</body>

</html> 