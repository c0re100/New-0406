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

<!--#include file="homecheck.asp"-->
<!--#include file="data2.asp"-->
<%
pagecount=15

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if

rs.open "select * from guestmessage where (send='"&info(0)&"' and receive='"&request.querystring("send")&"') or (send='"&request.querystring("send")&"' and receive='"&info(0)&"') order by id desc",conn,1,1

if rs.bof then
%>
<script language="Vbscript">
msgbox "沒有您需要的信息！",0,"提示"
history.back
</script>
<%
Response.End
end if
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
<font size="2" color="#E18C59"><b>留言歷史信息</b></font>
</div>
</td>
<td width="104">
<div align="center">

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
<div align="center">
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
<table border="1" width="695" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="476">□-您當前的位置----俠客寓所----留言本------留言歷史信息</td>
<td width="10"> </td>
<td class="p1" width="290">□-當前時間：<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td width="100%"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p1" width="400">□-給我的最新留言</td>
<td class="p1" width="78">□-<a href="javascript:history.back(1)"><font
color="#FFCC00">返回</font></a></td>
<td width="10"> </td>
<td class="p1" width="290"> </td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="0" cellpadding="0"
height="10" bordercolor="#E18C59">
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="100">
<%if dqpage>1 then%>
<form method="get" action="History222.asp">
<input type="hidden" name="page" value="1"><input
type="hidden" name="send"
value="<%=request.querystring("send")%>"><input type="submit"
value="首頁">
</form>
<%else%>
<input type="submit" value="首頁">
<%end if%>
</td>
<td class="p2" width="100">

<%if dqpage>1 then%>
<form method="get" action="History222.asp">
<input type="hidden" name="page" value="<%=dqpage-1%>"><input
type="hidden" name="send"
value="<%=request.querystring("send")%>"><input type="submit"
value="上頁">
</form>

<%else%>
<input type="submit" value="上頁">
<%end if%>

</td>
<td class="p2" width="100">

<%if dqpage<totalpage then%>
<form method="get" action="History222.asp">
<input type="hidden" name="page" value="<%=dqpage+1%>"><input
type="hidden" name="send"
value="<%=request.querystring("send")%>"><input type="submit"
value="下頁">
</form>

<%else%>
<input type="submit" value="下頁">
<%end if%>

</td>
<td class="p2" width="100">
<%if dqpage<totalpage then%>
<form method="get" action="History222.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="hidden" name="send"
value="<%=request.querystring("send")%>"><input type="submit"
value="尾頁">
</form>


<%else%>
<input type="submit" value="尾頁">
<%end if%>

</td>
<td class="p2" width="100">當前頁：<%=dqpage%></td>
<td class="p2" width="176">總頁：<%=totalpage%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="690" cellspacing="1" cellpadding="0"
bordercolor="#E18C59">
<tr>
<td class="p2" width="10"> </td>
<td class="p2" width="70">發送人</td>
<td class="p2" width="80">接收人</td>
<td class="p2" width="336">消息預覽</td>
<td class="p2" width="120">發送日期</td>
<td class="p2" width="60">是否刪除</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

%>
<tr>
<%message= replace(rs("message"),"<br>","")
if len(message)>24 then
message=left(message,24)&" ..."
end if%>
<td class="p3" width="10"> </td>
<td class="p3" width="70"><a
href="../friend/userinfo.asp?id=<%=rs("send")%>"><%=rs("send")%></a></td>
<td class="p3" width="80"><%=rs("receive")%></td>
<td class="p3" width="336"><a href="read.asp?id=<%=rs("id")%>"><%=message%></a></td>
<td class="p3" width="120"><%=rs("date")%> </td>
<td class="p3" width="60">
<form method="POST" action="delguest.asp">
<input type="hidden" name="id" value="<%=rs("id")%>"><input
type="submit" value="刪除">
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
<td class="p3" width="10"> </td>
<td class="p3" width="70"> </td>
<td class="p3" width="80"> </td>
<td class="p3" width="336"> </td>
<td class="p3" width="120"> </td>
<td class="p3" width="60"> </td>
</tr>
<%
next
end if
rs.close
conn.close
%>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="690" cellspacing="0" cellpadding="0">
<tr>
<td class="p1" width="100%">&nbsp;</td>
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

</html>