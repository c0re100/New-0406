<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>在線賭博</title>
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
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" text="#FFFFFF" link="#FFFFFF" topmargin="0" vlink="#FFFFFF" leftmargin="0"  background=bk.jpg>
<div align="center">

<table style=filter:shadow(color=#CC99FF)    border="1" cellspacing="0" cellpadding="1" width=95% bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center"  bgcolor=#006BA5>
<tr><td bgcolor="#EEFFEE" height="62" ><div align="center"><font color="#FF6633"> 在線賭博 </font></div></td></tr>
<tr><td>
<form  method=POST  action='duju/duju001.asp'><div align="center">
<br><input title="要求：戰鬥級大於30級、存款2000萬！" type=submit  name="銀兩局" value=銀兩局作莊>
<br><input title="要求：戰鬥級大於30級、法力5000點！" type=submit  name="法力局" value=法力局作莊>
<br><input title="要求：戰鬥級大於30級、泡點數3000點！" type=submit  name="泡點局" value=泡點局作莊>
<br><input title="要求：戰鬥級大於30級、金幣數4000個！" type=submit  name="金幣局" value=金幣局作莊>
<br><input title="吆喝其他玩家趕快下注！" type=submit  name="吆喝" value=各賭局吆喝>
</div>
</form>
</td></tr>
<tr><td><form  method=POST  action='duju/duju002.asp'><div align="center"><input title="銀兩賭局！" type=submit  name="銀兩局" value=銀兩局下注></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju003.asp'><div align="center"><input title="法力賭局！" type=submit  name="法力局" value=法力局下注></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju004.asp'><div align="center"><input title="泡點賭局！" type=submit  name="泡點局" value=泡點局下注></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju005.asp'><div align="center"><input title="金幣！" type=submit  name="金幣局" value=金幣局下注></div></form></td></tr>
<tr><td><form  method=POST  action='duju/DPK-XP.ASP'><div align="center"><input title="撲克牌局" type=submit  name="撲克牌局" value=撲克牌局></div></form></td></tr>
<tr><td><form  method=POST  action='duju/DMJ-XP.ASP'><div align="center"><input title="麻將牌局" type=submit  name="麻將牌局" value=麻將牌局></div></form></td></tr>
</table>

</div>
</body>
</html>
