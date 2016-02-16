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

<!--#include file="../data1.asp"-->

<%
rs.open "select * from bookmark_count order by count desc",conn,1,1
if rs.bof then
rs.close
%>
<script language="Vbscript">
msgbox"沒有開放的收藏夾！",0,"提示"
history.back
</script>
<%end if%>
<%
pagecount=20

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
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
<font size="2" color="#E18C59"><b>皇城書簽</b></font>
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
<td width="47" valign="top">
<div align="center">
<img src="../../jh/17.gif" width="47" height="251">
</div>
</td>
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
<table border="0" width="635" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">□-您當前的位置--俠客寓所---書簽中心&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;<a href="javascript:history.back(1)"><font color="#FFCC00">返回</font></a></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="40">
<%if dqpage>1 then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="1"><input
type="submit" value="首頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="首頁"
style="font-family: Tahoma; font-size: 12px">


<%end if%>
</td>
<td class="p2" width="40">
<%if dqpage>1 then%>

<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=dqpage-1%>"><input
type="submit" value="上頁"
style="font-family: Tahoma; font-size: 12px">
</form>
<%else%>
<input type="submit" value="上頁"
style="font-family: Tahoma; font-size: 12px">
<%end if%> </td>
<td class="p2" width="40">
<%if dqpage<totalpage then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=dqpage+1%>"><input
type="submit" value="下頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="下頁"
style="font-family: Tahoma; font-size: 12px">


<%end if%>
</td>
<td class="p2" width="40">
<%if dqpage<totalpage then%>
<form method="get" action="Marklist222.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="submit" value="尾頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="尾頁"
style="font-family: Tahoma; font-size: 12px">
<%end if%> </td>
<td class="p2" width="446">&nbsp;&nbsp;&nbsp; <a
href="useropen.asp"><font color="#FFCC00">書簽總排行</font></a></td>
<td class="p2" width="80">當前頁：<%=dqpage%>
</td>
<td class="p2" width="80">總頁：<%=totalpage%>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="100" align="center">排行名次</td>
