<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=440"%>
<!--#include file="dadata.asp"-->
<HTML><HEAD><TITLE>艶糃皊┍</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<LINK href="jd/the9.css" rel=stylesheet>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0
text=#000000 topMargin=0>
<TABLE bgColor=#000000 border=0 cellPadding=0 cellSpacing=0 height=24
width=760 align="center">
<TBODY>
<TR>
<TD width=48><IMG src="jd/empty.gif"></TD>
<TD width=34><IMG src="jd/position.gif" width="22" height="22"></TD>
<TD class=head-l width=660>眤讽玡竚<a href="../jh.asp"><font color="#FFFFFF">艶糃打</font></a>-&gt;<a href="jd.asp"><font color="#FFFFFF">艶糃皊┍</font></a>-&gt;ネΤ┋産穦</TD>
<TD align=right vAlign=top width=18><IMG
src="jd/rc.gif"></TD></TR></TBODY></TABLE>
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
<div align="center"><img src="jd/tubiao.gif" width="300" height="50"></div>
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
</td>
<td bgcolor=#faf4e9 valign=top width=39><img height=217
src="jd/index_xi.gif" width=39></td>
<td bgcolor=#faf4e9 valign=top width=295><br>
<table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
<tr bgcolor="#006633">
<td height="30" colspan="4" bgcolor="#FFFFFF">
<p align=center class="p9"><font color="#FF0000"><img src="jd/ssyx.gif" width="150" height="25"></font></p>
<tr bgcolor=#74E76D bordercolor="#000000">
<td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="25%">
<div align="center"><font size="2" color="#FF0000">産叫</font></div>
</td>
<td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="25%"><font size="2" color="#FF0000">産叫癸钩</font></td>
<td height="15" width="25%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">秨産丁</font></div>
</td>
<td height="15" width="25%" bgcolor="#66FF66" bordercolor="#FF6600">
<div align="center"><font size="2" color="#FF0000">畊</font></div>
</td>
</tr>
<!--#include file="dadata.asp"--> <%
sql="SELECT id,局Τ,把,丁 FROM 産穦 where  産穦='ネΤ┋'"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor=#DEAD63>
<td width="15%" bgcolor="#FFFFFF" class="calen-curr"><font size="2">
<center>
<%=rs("局Τ")%>
</center>
</font> <font size="2"></font>
<td width="5%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"><%=rs("把")%></font>
<td width="45%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"><%=rs("丁")%>
</font></td>
<td width="35%" bgcolor="#FFFFFF" class="calen-curr"><font size="2">
<center>
<font size="2">
<%if info(0)=rs("局Τ") or info(0)=rs("把") then %>
<a href=ruxi1.asp?id=<%=rs("id")%>><span class="calen-curr">畊</span></a>
<%end if%>
</font>
</center>
</font></td>
</tr>
<%
rs.movenext
loop
connt.close
set rs=nothing
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
<div style="HEIGHT: 50px; WIDTH: 150px"> <marquee direction=up height=150 onMouseOut=this.start()
onMouseOver=this.stop() scrollamount=1 scrolldelay=30>艶糃皊┍い<br>
﹚皊畊叫克狟ね<br>
ㄓ硂把﹚皊畊皊畊だ挡盉産ネら産<br>
叫産叫竜産单单ぃ皊畊把┮眔<br>
眔狥﹁琌ぃ讽礛把计琌ぃ妓 <font
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
