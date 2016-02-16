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


<!--#include file="data1.asp"-->
<html>

<head>
<title>靈劍江湖-寵物商店</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
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

<body bgcolor="#000000" text="#FFFFFF">

<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
<td width="744" background="../../jh/tiao20.gif" height="83">
<table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"><span class="zilong"><font color="#FFCC00">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %>
</font></span></font></font></font></td>
<td width="475" height="26">
<div align="center">
<font size="2" color="#E18C59"><b>留言板</b></font>
</div>
</td>
<td width="104">
<div align="center">
<a href="../../jh.asp" target="_top"><font color="#E18C59">返回皇城中心</font></a>
</div>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
</td>
</tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="17" background="../../jh/tiao10.gif">&nbsp;</td>
<td valign="top" width="639">
<div align="center" style="width: 695; height: 470">
<div align="center">
<center>
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">
</center>
</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="695" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="552" height="27">您的位置：俠客寓所&gt;&gt;留言板
當前時間：<%=date%><%=time%>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="475" cellspacing="1" cellpadding="0"
height="50" bordercolor="#E18C59">
<tr>
<td class="p2" width="469" height="16">提示：</td>
</tr>
<tr>
<td class="p3" width="469" height="5">&nbsp;歡迎使用您的留言版，這裡您可以給江湖中有家的俠士留言。當您在留言後選擇了</td>
</tr>
<tr>
<td class="p3" width="469" height="8">&nbsp;需要回執，當您的朋友讀了您的留言後，系統會自動發給您提示信息。在這裡您</td>
</tr>
<tr>
<td class="p3" width="469" height="9">&nbsp;還可以將你的朋友加到好友清單中，把你不喜歡或者惡意搗亂的人加到黑名單中。</td>
</tr>
<tr>
<td class="p3" width="469" height="5">&nbsp;</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="556" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="267">□-給我的最新留言</td>
<td width="2"></td>
<td class="p1" width="289">□-我想給我的朋友留言</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="609" cellspacing="0" cellpadding="0">
<tr>
<td width="335" valign="top">
<table border="1" width="312" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="91">□-發言內容</td>
<td class="p2" width="65">□-發言人</td>
<td class="p2" width="78">□-發言時間</td>
<td class="p2" width="54">狀態</td>
<td class="p2" width="40">歷史</td>
</tr>
<%
rs.open "select * from guestmessage where receive='"&info(0)&"' and messagetype='一般消息' order by state desc",conn,1,1
count=0
do while not rs.eof and count<8
%>
<% message=replace(rs("message"),"<br>","")
if len(message)>14 then
message=left(message,14)&" ..."
end if%>
<tr>
<td class="p3" width="91">◎&nbsp;<a
href="read.asp?id=<%=rs("id")%>"><%=message%></a></td>
<td class="p3" width="65"><a
href="../friend/userinfo.asp?id=<%=rs("send")%>"><%=rs("send")%></a></td>
<td class="p3" width="78"><%=left(rs("date"),10)%></td>
<td class="p3" width="54"><%
if rs("state")="1" then
Response.Write( "<b>未讀</b>")
else
Response.Write( "已讀")
end if
%></td>
<td class="p3" width="40"><a
href="history.asp?send=<%=rs("send")%>"><font
color="#FFCC00">點擊</font></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<8 then
for i=1 to 8-count
%>

<tr>
<td class="p3" width="91"> </td>
<td class="p3" width="65"> </td>
<td class="p3" width="78"> </td>
<td class="p3" width="54"> </td>
<td class="p3" width="40"> </td>
</tr>
<%
next
end if
rs.close
%>

<%
pagecount=5

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if

searchsql="select * from guestmessage where receive='"&info(0)&"' and (messagetype='系統消息' or messagetype='讀取確認') order by state desc"

rs.open searchsql,conn,1,1

temp1=rs.recordcount/pagecount
if temp1-int(rs.recordcount/pagecount)<>0 then
totalpage=int(temp1+1)
else
totalpage=temp1
end if
if dqpage>totalpage then
dqpage=1
end if

if dqpage<>1 then
temp2=(dqpage-1)*pagecount
rs.move temp2

end if
%>

<tr>
<td class="p2" width="328" colspan="5" align="right"><a
href="disp_guest.asp"><font color="#FFCC00">更多留言</font>...</a></td>
</tr>
</table>
<table border="1" width="325" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="358" colspan="7" bgcolor="#FFCB31">□-系統信息</td>
</tr>
<tr>
<td class="p2" width="2"></td>
<td class="p2" width="54">


<%if dqpage>1 then%>

<form method="get" action="Index222.asp">
<input type="hidden" name="page" value="1"><input
type="submit" value="首頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>

<input type="submit" value="首頁"
style="font-family: Tahoma; font-size: 12px">

<%end if%>



</td>
<td class="p2" width="47">

<%if dqpage>1 then%>
<form method="get" action="Index222.asp">
<input type="hidden" name="page" value="<%=dqpage-1%>"><input
type="submit" value="上頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="上頁"
style="font-family: Tahoma; font-size: 12px">

<%end if%>

</td>
<td class="p2" width="47">


<%if dqpage<totalpage then%>

<form method="get" action="Index222.asp">
<input type="hidden" name="page" value="<%=dqpage+1%>"><input
type="submit" value="下頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="下頁"
style="font-family: Tahoma; font-size: 12px">
<%end if%>

</td>
<td class="p2" width="87">

<%if dqpage<totalpage then%>
<form method="get" action="Index222.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="submit" value="尾頁"
style="font-family: Tahoma; font-size: 12px">
</form>



<%else%>
<input type="submit" value="尾頁">
<%end if%>


</td>
<td class="p2" width="68">當前頁：<%=dqpage%></td>
<td class="p2" width="46">總頁：<%=totalpage%></td>
</tr>
</table>
<table border="1" width="317" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="62">□-內容預覽</td>
<td class="p2" width="49">發信人</td>
<td class="p2" width="80">發信時間</td>
<td class="p2" width="33">狀態</td>
<td class="p2" width="48">歷史</td>
<td class="p2" width="45">刪除</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount
message=replace(rs("message"),"<br>","")
if len(message)>14 then
message=left(message,14)&" ..."
end if
%>

<tr>
<td class="p3" width="62">◎&nbsp;<a
href="read.asp?id=<%=rs("id")%>"><%=message%></a></td>
<td class="p3" width="49"><%=rs("send")%> </td>
<td class="p3" width="80"><%=left(rs("date"),10)%> </td>
<td class="p3" width="33">
<%
if rs("state")="1" then
Response.Write( "<b>未讀</b>")
else
Response.Write( "已讀")
end if
%>
</td>
<td class="p3" width="48"><a
href="history.asp?send=<%=rs("send")%>"><font
color="#FFCC00">點擊</font></a></td>
<td class="p3" width="45">
<form method="POST" action="delguest.asp">
<input type="hidden" name="id" value="<%=rs("id")%>"><input
type="submit" value="刪除"
style="font-family: Tahoma; font-size: 12px">
</form>
</td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count
%>
<tr>
<td class="p3" width="62" height="30"></td>
<td class="p3" width="49" height="30"></td>
<td class="p3" width="80" height="30"></td>
<td class="p3" width="33" height="30"></td>
<td class="p3" width="48" height="30"></td>
<td class="p3" width="45" height="30"></td>
</tr>
<%
next
end if
rs.close

%>
</table>
</td>
<td width="11"></td>
<td width="320" valign="top">
<table border="1" width="173" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<form method="POST" action="post.asp">
<td class="p3" width="290" colspan="2">□-我朋友的名字：<input
type="text" name="name" size="17" maxlength="50"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
<tr>
<td class="p3" width="290" colspan="2">□-留言內容：</td>
</tr>
<tr>
<td width="290" colspan="2"><textarea rows="4"
name="message" cols="41"
style="font-family: Tahoma; font-size: 12px"></textarea></td>
</tr>
<tr>
<td class="p3" width="144">&nbsp;<input type="checkbox"
name="receipt" value="ON"> 是否需要回執？</td>
<td class="p3" width="150"><input type="submit"
value="提交" name="B1"
style="font-family: Tahoma; font-size: 12px"><input
type="reset" value="全部重寫" name="B2"></td>
</tr>
</form>
</table>
<%
rs.open "select * from usertype where user='"&info(0)&"' and type='友' order by id desc"
friendcount=rs.recordcount
%>
<table border="1" width="142" align="left" cellspacing="1"
cellpadding="0" bordercolor="#E18C59">
<tr>
<td class="p2" width="100%">我的朋友</td>
</tr>
<%
count=0
do while not rs.eof and count<12
%>
<tr>
<td class="p3" width="100%"><a
href="send.asp?name=<%=rs("name")%>"><%=rs("name")%></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<12 then
for i=1 to 12-count
%>
<tr>
<td class="p3" width="100%">&nbsp;</td>
</tr>
<%
next
end if
rs.close
%>
<tr>
<td class="p2" width="100%"><%if friendcount>0 then%>
<a href="../friend/hylb.asp"><font color="#FFCC00">管理我的朋友</font></a><font
color="#FFCC00"> </font>
<%end if%>
</td>
</tr>
</table>
<%
rs.open "select * from usertype where user='"&info(0)&"' and type='黑' order by id desc"
blackcount=rs.recordcount
%>

<table border="1" width="124" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="118">黑名單</td>
</tr>
<%
count=0
do while not rs.eof and count<12
%>

<tr>
<td class="p3" width="118"><a
href="send.asp?name=<%=rs("name")%>"><%=rs("name")%></a></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<12 then
for i=1 to 12-count
%>

<tr>
<td class="p3" width="118">&nbsp;</td>
</tr>
<%
next
end if
rs.close
%>
<tr>
<td class="p2" width="118"><%if blackcount>0 then%>
<a href="../friend/black.asp"><font color="#FFCB31">管理我的黑名單</font></a>
<%end if%> </td>
</tr>
</table>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="647" cellspacing="0" cellpadding="0">
<tr>
<td class="p1" width="645">&nbsp;</td>
</tr>
</table>
</center>
</div>
</div>
</td>
<td width="25" background="../../jh/tiao10.gif"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<div align="right">
<img src="../../jh/tiao19.gif" width="728" height="31">
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>
<%conn.close%>

</html>

<html><script language="JavaScript">                                                                  </script></html>
