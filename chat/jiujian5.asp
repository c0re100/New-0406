<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
 
<%if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
nickname=info(0)
who=Trim(Request.QueryString("who"))
if who="" then who=nickname
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
chatroombgimage=ljjh_chatroombgimage
chatroombgcolor=ljjh_chatroombgcolor
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,門派 from 用戶 where id="&info(9),conn
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖人物！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
else
pai=rs("門派")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>破鞭式 </title>
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
<tr bgcolor=ffff00><td>
破鞭式<br>
戰鬥等級大於5級<br>
花費銀兩5025000兩 <br>
攻擊對像等級小於10級 <br>
</td></tr>
<tr><td>
<div align="center">
<form method=POST action='jiujian5ok.asp'>
攻擊對像<br>
<select name="to1">
<%
for i=0 to x
if show(i)<>info(0) then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
<br>
<input type=submit value="施展絕技" name=submit>
<br>
</form>
<form method=POST action='jiujian5ok.asp'>
<select name="to1">
<%
if session("nowinroom") = 1 then
for gw=5 to 20
gw=gw+3
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=20 to 150
gw=gw+15
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=180 to 390
gw=gw+30
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=490 to 780
gw=gw+50
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=801 to 1100
gw=gw+49
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
    </select>
    <br>
    <br>
    <input type=submit value="施展絕技" name=submit2>
  </form>
</td></tr>
</table></TD>
<TD background="./fq/tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="./fq/tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a5.gif" width="90" height="10"></TD>
<TD width="57" height="10"><IMG src="./fq/tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML>
<%end if%>