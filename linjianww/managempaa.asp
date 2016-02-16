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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
id=Request.QueryString("ID")
sql="Select * from 職業技能 where 技能='"&Id&"'"
rs=conn.execute(sql)
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
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
<body bgcolor="#FAF0E2" text="#000000" link="#000080" alink="#800000" vlink="#000080"  background="../ff.jpg">
<form action="updatmpaa.asp" method=POST>
<table border=1 cellspacing=0 cellpadding=3 align="center" width="500" bordercolordark="#FFFFFF">
<tr><td>技能</td><td><input name="aa" size=40 maxlength=50 value="<%=RS("技能")%>"></td></tr>
<tr><td>修練條件</td><td><input name="cc" value="<%=RS("修練條件")%>" size=40></td></tr>
<tr><td>修練法力</td><td><input name="dd" value="<%=RS("修練法力")%>" size=40 maxlength=50></td></tr>
<tr><td>消耗法力</td><td><input name="ee" value="<%=RS("消耗法力")%>" size=40 maxlength=50></td></tr>
<tr><td>基本攻擊</td><td><input name="ff" value="<%=RS("基本攻擊")%>" size=40 maxlength=50></td></tr>
<tr><td>魔法</td><td><input name="gg" value="<%=RS("魔法")%>" size=40 maxlength=50></td></tr>
<tr><td>圖檔</td><td><input name="hh" value="<%=RS("圖檔")%>" size=40></td></tr>
<tr><td>施法說明</td><td><input name="ii" value="<%=RS("施法說明")%>" size=40></td></tr>
</table>
<input type="HIDDEN" name="action" value="RegSubmit">
<input type="SUBMIT" name=okcc value="更新招式">&nbsp;&nbsp;<input type="SUBMIT" name=okcc value="新增招式">
</div>
</form>
</body>
</html>
<%
set rs=nothing 
%>
