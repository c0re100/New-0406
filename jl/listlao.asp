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
<title>在押人犯</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background=../bk.jpg link="#0000FF" vlink="#0000FF" alink="#0000FF">
<br>
<table border="1" cellspacing="0" cellpadding="0" width="560" align="center">
  <tr> 
    <td width="15%" nowrap bgcolor="#99CCFF"> 
      <div align="center"> 犯人 </div>
    <td width="15%" nowrap bgcolor="#99CCFF"> 
      <div align="center"> 門派 </div>
    </td>
    <td width="11%" bgcolor="#99CCFF"> 
      <div align="center"> 身份 </div>
    </td>
    <td width="27%" bgcolor="#99CCFF"> 
      <div align="center"> 狀 態</div>
    </td>
    <td colspan="2" bgcolor="#99CCFF"> 
      <div align="center">操 作 </div>
    </td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,姓名,門派,身份,狀態,登錄 FROM 用戶 where 狀態='獄' or 狀態='牢'",conn
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#FFCC99"  onMouseOut="this.bgColor='#FFCC99';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td width="15%" height="31"><%=rs("姓名")%> 
    <td width="15%" height="31"><%=rs("門派")%></td>
    <td width="11%" nowrap height="31"><%=rs("身份")%></td>
    <%if rs("狀態")="獄" then %>
    <td width="27%" height="31"> 
      <div align="center"> <a href="shifang.asp?name=<%=rs("姓名")%>"><font color="#0000FF">交錢100萬兩釋放</font></a>|<a href="jieyu.asp?name=<%=rs("姓名")%>"
title="成功率只有30%"><font color="#FF0000">劫獄</font></a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>"><font color="#FF0000">釋放</font></a> 
        <%end if%>
      </div>
    </td>
    <% else
if rs("登錄")>now() then%>
    <td width="23%" height="31"> 
      <div align="center"> <font color="#0000FF">無權操作</font> 
        <%if  info(0)="江湖行" then%>
        <a href="shifang.asp?name=<%=rs("姓名")%>"><font color="#0000FF">交錢50000兩釋放</font></a> 
        <%else%>
        <font color="#0000FF">無權操作</font> 
        <%end if%>
        <font color="#FF0033">不能保釋</font>|釋放時間：<%=rs("登錄")%><a href="yongka.asp?name=<%=rs("姓名")%>"><font color="#0000FF">用卡</font></a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>"><font color="#FF0000">釋放</font></a> 
        <%end if%>
        |<a href="jieyu.asp?name=<%=rs("姓名")%>"
title="成功率只有30%"><font color="#FF0000">劫獄</font></a> </div>
    </td>
    <%else%>
    <td width="9%" height="31"> 
      <div align="center"> <font color="#FF0033">已釋放</font> </div>
    </td>
    <% conn.execute("update 用戶 set 狀態='正常',登錄=now() where 姓名='"&rs("姓名")&"'")
end if%>
    <%end if%>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p align="center"><a href="#" onClick="window.close()">關 閉</a></p>
</body>

</html>

<html></html> 



