<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
filname=Session("ljjh_filname")
nickname=info(0)
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if nickname="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&nickname&" ")=0 then Response.Redirect "chaterr.asp?id=001"
show=Split(Trim(useronlinename)," ",-1)
j=UBound(show)
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>屏蔽某人的話</title>
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
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" topmargin="0"  background=bk.jpg>
<div align="center">
<tr>
<td>
<table border="1" cellspacing="0" bgcolor="#EEFFEE" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" width="154" cellpadding="3" height="267">
<form method="post" action="filterok.asp" name="">
<tr align="center">
<td><font color="red">選擇屏蔽對像</font>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
<td><%for i=0 to j
if CStr(show(i))<>CStr(nickname) then%><input type="checkbox" name="filtername" value="<%=show(i)%>"<%if Instr(filname," "&show(i)&",")<>0 then Response.Write " checked"%>><%=show(i)%><br>
<%end if
next%></td>
</tr>
</table>
<table width="100%" border="0">
<tr align="center">
<td>
<input type="submit" name="ok" value=確定><input type="submit" name="clearall" value=取消>
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</td>
</tr>
</div>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body></html> 