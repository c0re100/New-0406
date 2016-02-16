<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<title>技能管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
td, div, form, option, p, td, br{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{border-color:#000000; border-width:1px; padding:1px; BACKGROUND: #ffffff url('2_68.gif');FONT-SIZE: 9pt; HEIGHT: 18px; }
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
<style type="text/css">
.aq {  font-size: 9pt; line-height: 150%; color: #efefef; filter: DropShadow(Color=#efefef, OffX=1, OffY=1, Positive=1)}
</style>
</head>
<body bgcolor="#FFFFFF" background="../ff.jpg">
<font size="2">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("ljjh_usermdb")
conn.open connstr
set subj=server.createobject("adodb.recordset") 
set owner=server.createobject("adodb.recordset")
subj.open "Select * From 職業技能",conn,0,1
if not subj.eof then 
%></font> 
<table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<tr bgcolor="#CCCCCC" bordercolor="#CCCCCC"> 
<td align="center" bgcolor="#669900">&nbsp;<font color="#FFFFFF">打架招式</font></td>
<td align="center" bgcolor="#669900"><font color="#FFFFFF">修練條件</font></td>
<td align="center" bgcolor="#669900"><font color="#FFFFFF">修練法力</font></td>
<td align="center" bgcolor="#669900"><font color="#FFFFFF">消耗法力</font></td>
<td align="center" bgcolor="#669900"><font color="#FFFFFF">基本攻擊</font></td>
<td align="center" bgcolor="#669900"><font color="#FFFFFF">魔法</font></td>
</tr>
<%do while not subj.eof%> 
<tr bgcolor="#FFFFFF">
<td><div align="center"><%=subj("技能")%></div></td>
<td><div align="center"><%=subj("修練條件")%></div></td>
<td><div align="center"><%=subj("修練法力")%></div></td>
<td><div align="center"><%=subj("消耗法力")%></div></td>
<td><div align="center"><%=subj("基本攻擊")%></div></td>
<td><div align="center"><%=subj("魔法")%></div></td></tr>
<tr><td><div align="center"><%=subj("圖檔")%></div></td>
<td colspan="4">施法說明:<%=subj("施法說明")%></td>
<td><p align="center"><a href="managempaa.asp?id=<%=SUBJ("技能")%>">修改</a>&nbsp;<a href="delaa.asp?id=<%=SUBJ("id")%>">刪</a></td></tr>
<%
subj.movenext
loop
else
%> 
</table>
<table align="center" width="450" cellpadding="20" cellspacing="50" border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<td height="14"> 
<div align="center"><font color="#FFFFFF">還沒有一個技能呢！</font> 
<% subj.close 
set subj=nothing 
end if%>
</div>
</table>
<div align="center"></div>
<p align="center">[<a href="adminaa.asp">====更新====</a>]
</body>