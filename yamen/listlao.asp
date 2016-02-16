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
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#006666" background="../ff.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0">
<br>
<table border=0 cellspacing="2" cellpadding="2" width="560" align="center">
  <tr bgcolor="#CC0000"> 
    <td width="15%"> 
      <div align="center"><font color="#FFFFFF"> 犯人 </font></div></td>
    <td width=15%> 
      <div align="center"><font color="#FFFFFF"> 門派 </font></div></td>
    <td width=11%> 
      <div align="center"><font color="#FFFFFF"> 身份 </font></div></td>
    <td width=27%> 
      <div align="center"><font color="#FFFFFF"> 狀 態 </font></div></td>
    <td colspan=2> 
      <div align="center"><font color="#FFFFFF">操 作 </font></div></td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 姓名,門派,身份,狀態,id,登錄 FROM 用戶 where 狀態='獄' or 狀態='牢'",conn
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#330000"> 
    <td><font color="#FFFFCC"><%=rs("姓名")%></td>
    <td><font color="#FFFFCC"><%=rs("門派")%></td>
    <td><%=rs("身份")%></td>
    <%if rs("狀態")="獄" then %>
    <td width="27%"><div align="center"> <a href="shifang.asp?name=<%=rs("姓名")%>">交錢100萬兩釋放</a>|<a href="jieyu.asp?name=<%=rs("姓名")%>" title="成功率只有30%">劫獄</a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>">釋放</a> 
        <%end if%>
      </div></td>
    <% else
if rs("登錄")>now() then%>
    <td width="23%" height="31"> <div align="center"> <font color="#FFFFCC">無權操作</font> 
        <%if info(0)="江湖總站" then%>
        <a href="shifang.asp?name=<%=rs("姓名")%>">交錢50000兩釋放</a> 
        <%else%>
        無權操作 
        <%end if%>
        不能保釋|釋放時間：<%=rs("登錄")%><a href="yongka.asp?name=<%=rs("姓名")%>">用卡</a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>">釋放</a> 
        <%end if%>
        |<a href="jieyu.asp?name=<%=rs("姓名")%>"
title="成功率只有30%">劫獄</a></div></td>
    <%else%>
    <td width="9%"> <div align="center">已釋放</div></td>
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
