<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="dadata.asp"-->
<%if not IsArray(Session("info")) then Response.Redirect "error.asp?id=440"
info=Session("info")
%>
<HTML><head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<TITLE>靈劍大酒店</TITLE>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</HEAD>
<BODY bgColor=#FF0000 leftMargin=0 text=#000000 topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
<TBODY>
<TR>
<TD width=236><IMG height=119 src="jd/ggg.gif" width=236></TD>
<TD background=jd/index_bg1.gif bgColor=#d70000 vAlign=top>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=524>
<TBODY>
<TR>
<TD height=8 width=120></TD>
<TD width=62></TD>
<TD width=47></TD>
<TD width=100></TD>
<TD width=74></TD>
<TD width=42></TD>
<TD width=79></TD>
</TR>
<TR>
<TD colspan="7">
<div align="center"><img src="jd/tubiao.gif"></div>
</TD>
</TR>
</TBODY>
</TABLE>
</TD></TR></TBODY></TABLE>
<table border=0 cellpadding=0 cellspacing=0 width=760 align="center">
<tbody>
<tr>
<td bgcolor=#d70000 height=379 valign=top width=180>
<table bgcolor=#d70000 border=0 cellpadding=0 cellspacing=0 width=180>
<tbody>
<tr>
<td><img height=60 src="jd/index_t2.gif" width=149></td>
</tr>
<tr>
<td><a href=yiyi.asp
><img border=0 height=26
src="jd/index_left1.gif" width=149></a></td>
</tr>
<tr>
<td><a href=byby.asp
><img border=0 height=20
src="jd/index_left2.gif" width=149></a></td>
</tr>
<tr>
<td><a href=ssss.asp
><img border=0 height=19
src="jd/index_left3.gif" width=149></a></td>
</tr>
<tr>
<td><a href="smsm.asp"
><img border=0 height=21
src="jd/index_left4.gif" width=149></a></td>
</tr>
<tr>
<td><a href=wcwc.asp
><img border=0 height=20
src="jd/index_left5.gif" width=149></a></td>
</tr>
<tr>
<td><a href=lglg.asp
><img border=0 height=19
src="jd/index_left6.gif" width=149></a></td>
</tr>
<tr>
<td><a href=lqlq.asp
><img border=0 height=22
src="jd/index_left7.gif" width=149></a></td>
</tr>
<tr>
<td><a href=btbt.asp
><img border=0 height=20
src="jd/index_left8.gif" width=149></a></td>
</tr>
<tr>
<td><a href=tctc.asp
><img border=0 height=20
src="jd/index_left9.gif" width=149></a></td>
</tr>
<tr>
<td><img height=135 src="jd/index_t3.gif"
width=180></td>
</tr>
</tbody>
</table>
<div id="Layer1" style="position:absolute; width:105px; height:51px; z-index:1; left: 10px; top: 372px"><b><font color="#FFFF00">喫酒點這裡！</font></b></div>
</td>
<td bgcolor=#faf4e9 valign=top width=39><img height=217
src="jd/index_xi.gif" width=39></td>
<td bgcolor=#faf4e9 valign=top width=295><br>
<table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
<tr bgcolor="#006633">
<td height="30" colspan="4" bgcolor="#FFFFFF">
<p align=center class="p9"><font color="#FF0000"><img src="jd/djx.gif" width="150" height="25"></font></p>

<tr bgcolor=#74E76D bordercolor="#000000">
<td height="15" width="25%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">宴會名</font></div>
</td>
<td height="15" width="33%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">規格</font></div>
</td>
<td height="15" width="24%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">售價(兩)</font></div>
</td>
<td height="15" width="18%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">訂購</font></div>
</td>
</tr>
<!--#include file="dadata.asp"--> <%
sql="SELECT 宴會名,數量,售價,id FROM 宴會"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor=#DEAD63>
<td width="25%" bgcolor="#FFFFFF" class="calen-curr"><font size="2">
<center>
<%=rs("宴會名")%>
</center>
</font>
<td width="33%" bgcolor="#FFFFFF" class="ct-def4"><font size="2">
<% if rs("宴會名")="一心一意" then %><font color=blue>同門</font>
<%else
if rs("宴會名")="比翼雙飛" then %><font color=blue>情人</font>
<%else
if rs("宴會名")="三生有幸" then %><font color=blue>朋友</font>
<%else
if rs("宴會名")="海誓山盟" then %><font color=blue>夫妻</font>
<%else%><font color=blue>所有</font>
<%end if%>
<%end if%>
<%end if%>
<%end if%><%=rs("數量")%><span class="calen-curr">人參加</span></font></td>
<td width="24%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"><%=rs("售價")%></font></td>
<td width="18%" bgcolor="#FFFFFF" class="calen-curr"><font size="2">
<center>
<font size="2">
<%if rs("宴會名")="比翼雙飛" then %><a href="#" onclick="window.open('jiudian3.asp?id=<%=rs("id")%>','yanqing','scrollbars=no,resizable=no,width=395,height=223')" title="向別人求婚，先來這裡宴請他(她)呀。">
<%else
if rs("宴會名")="三生有幸" then %><a href="#" onclick="window.open('jiudian4.asp?id=<%=rs("id")%>','yanqing','scrollbars=no,resizable=no,width=395,height=223')" title="宴請你的朋友一人。">
<%else
if rs("宴會名")="海誓山盟" then %><a href=jiudian5.asp?id=<%=rs("id")%> title="為你的愛人和自己定一份酒席!">
<%else
if rs("宴會名")="一心一意" and info(5)<>"無" and info(5)<>"未定" then %><a href=jiudian2.asp?id=<%=rs("id")%> title="為自己的同門兄弟定酒宴!">
<%else%>
<a href=jiudian1.asp?id=<%=rs("id")%> title="宴請所有的人!">
<%end if%><%end if%><%end if%><%end if%>
<span class="calen-curr">定酒席</span></a></font>
</center>
</font></td>
</tr>
<%
rs.movenext
loop
rs.close
connt.close
set rs=nothing
set conn=nothing
%>
</table>
<br>
<img
border=0 height=44 name=Image90 src="jd/index_r1_c16.gif"
width=123> </td>
<td background=jd/index_bg2.gif bgcolor=#faf4e9 width=223>
<table border=0 cellpadding=10 cellspacing=15 class=ct-def2 width=185>
<tbody>
<tr bgcolor=#ffffff>
<td height=150><img height=1 src="jd/empty.gif" width=1>
<div style="HEIGHT: 50px; WIDTH: 150px"> <span id="adver" style="visibility:hidden;position:absolute;top:0;width:50;left=0"><script>
var tc_user="aresky";
var tc_class="2";
var tc_union="";
var tc_type="1";
</script>
</span><marquee direction=up height=150 onMouseOut=this.start()
onMouseOver=this.stop() scrollamount=1 scrolldelay=30>在靈劍大酒店中，你可以定酒席，請你的親朋好友來這參加你定下的酒席，酒席分為結婚宴，生日宴請功宴，請罪宴等等。不同的酒席，參加者所得到得東西是不同的，當然參加的人數也是不一樣的。
<font
color=#cc0000></font><br>
</marquee></div>
</td>
</tr>
</tbody>
</table>
</td>
<td bgcolor=#d70000 width=23>&nbsp;</td>
</tr>
<tr>
<td bgcolor=#d70000 width=180><img height=44 src="jd/index_t4.gif"
width=149></td>
<td bgcolor=#d70000 width=39>&nbsp;</td>
<td bgcolor=#d70000 width=295>&nbsp;</td>
<td bgcolor=#d70000 width=223>&nbsp;</td>
<td bgcolor=#d70000 width=23>&nbsp;</td>
</tr>
</tbody>
</table>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
<TBODY>
<TR>
<TD bgColor=#ff700d>&nbsp;</TD></TR></TBODY></TABLE>
</BODY></HTML>
