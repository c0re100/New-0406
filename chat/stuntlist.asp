<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>特技列表</title>
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
<HTML>
<HEAD>
<TITLE>法術</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY bgcolor="#ffffff" background="bk.jpg">
<TABLE border="0" cellpadding="0" cellspacing="0">
<TR> 
<TD align="center" height="10" width="20" valign="top"><IMG src="./fq/tb-a1.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a6.gif" height="10" width="90"></TD>
<TD height="10" align="center" valign="top" width="57"><IMG src="./fq/tb-a7.gif" width="20" height="20" border="0"></TD>
</TR>
<TR> 
<TD width="20" background="./fq/tb-a2.gif" height="80"></TD>
<TD width="90" height="80" align="center" valign="top" background="./fq/tb-a3.gif">
<table border="0" cellspacing="0" cellpadding="1" width="90" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="287" >
<div align="center">
<tr align=center><td><form  method=POST  action='stunt.asp'><input type=submit  name="確定" value="㊣連續技㊣"></form></td></tr>
<tr align=center><td><form  method=POST  action='compact.asp'><input type=submit  name="確定" value="→合體技←"></form></td></tr>
<tr align=center><td><form  method=POST  action='pet.asp'><input type=submit  name="確定" value="※寵物技※"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian1.asp'><input type=submit  name="確定" value="☆總訣式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian2.asp'><input type=submit  name="確定" value="☆破刀式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian3.asp'><input type=submit  name="確定" value="☆破槍式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian4.asp'><input type=submit  name="確定" value="☆破劍式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian5.asp'><input type=submit  name="確定" value="☆破鞭式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian6.asp'><input type=submit  name="確定" value="☆破索式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian7.asp'><input type=submit  name="確定" value="☆破箭式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian8.asp'><input type=submit  name="確定" value="☆破掌式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='jiujian9.asp'><input type=submit  name="確定" value="☆破氣式☆"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao1.asp'><input type=submit  name="確定" value="分筋錯骨"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao2.asp'><input type=submit  name="確定" value="化骨綿掌"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao3.asp'><input type=submit  name="確定" value="明月神功"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao4.asp'><input type=submit  name="確定" value="排雲掌法"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao5.asp'><input type=submit  name="確定" value="浮影劍法"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao6.asp'><input type=submit  name="確定" value="雷霆三式"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao7.asp'><input type=submit  name="確定" value="覆雨翻雲"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao8.asp'><input type=submit  name="確定" value="冰雪飛塵"></form></td></tr>
<tr align=center><td><form  method=POST  action='juezhao9.asp'><input type=submit  name="確定" value="天外飛仙"></form></td></tr>
</table></TD>
<TD background="./fq/tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="./fq/tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a5.gif" width="90" height="10"></TD>
<TD width="57" height="10"><IMG src="./fq/tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML>