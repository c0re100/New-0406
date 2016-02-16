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

 
<!--#include file="../data2.asp"-->
<%
if info(0)="" then
%>
<script language="Vbscript">
msgbox"非法入侵！",0,"警告"
history.back
</script>
<%
Response.End
end if%>
<%
if not isnumeric(Request.QueryString("page")) then
%>
<script language="Vbscript">
msgbox"輸入不是數字！",0,"警告"
history.back
</script>
<%
Response.End
end if%>

<%
pagecount=10

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if
if trim(request.querystring("name"))<>"" then
rs.open "select * from diarymessage where open='是' and instr(user,'"&request.querystring("name")&"') order by id desc ",conn,1,1
else
rs.open "select * from diarymessage where open='是' order by id desc ",conn,1,1
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
<font size="2" color="#E18C59"> <b>日記中心</b></font>
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
<td class="p1" width="334">□-您當前的位置-<a
href="../index.asp"><font color="#FFCC00">俠客寓所</font></a>--日記中心</td>
<td class="p1" width="440">□-當前時間：<%=date%>&nbsp;&nbsp;<%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="638" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="10"></td>
<td class="p1" width="714">□-日記中心</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="649" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="10"></td>
<td class="p2" width="40"><%if dqpage>1 then%>
<form method="get" action="disp_diary222.asp">
<input type="hidden" name="page" value="1"><input
type="hidden" name="name"
value="<%=request.querystring("name")%>"><input type="submit"
value="首頁">
</form>
<%else%>
<input type="submit" value="首頁">
<%end if%> </td>
<td class="p2" width="40"><%if dqpage>1 then%>
<form method="get" action="disp_diary222.asp">
<input type="hidden" name="page" value="<%=dqpage-1%>"><input
type="hidden" name="name"
value="<%=request.querystring("name")%>"><input type="submit"
value="上頁">
</form>
<%else%>
<input type="submit" value="上頁">
<%end if%> </td>
<td class="p2" width="40"><%if dqpage<totalpage then%>
<form method="get" action="disp_diary222.asp">
<input type="hidden" name="page" value="<%=dqpage+1%>"><input
type="hidden" name="name"
value="<%=request.querystring("name")%>"><input type="submit"
value="下頁">
</form>

<%else%>
<input type="submit" value="下頁">
<%end if%> </td>
<td class="p2" width="40"><%if dqpage<totalpage then%>
<form method="get" action="disp_diary222.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="hidden" name="name"
value="<%=request.querystring("name")%>"><input type="submit"
value="尾頁">
</form>

<%else%>
<input type="submit" value="尾頁">
<%end if%></td>
<td class="p2" width="260">
<form method="get" action="disp_diary222.asp">
&nbsp;按照姓名查詢：<input type="text" name="name"
size="13" maxlength="5"><input type="submit" value="查詢">
</form>
</td>
<td class="p2" width="80">當前頁：<%=dqpage%></td>
<td class="p2" width="80">總頁：<%=totalpage%></td>
<td class="p2" width="145">
<form method="get" action="disp_diary222.asp">
轉到：<input type="text" name="page" size="10"
maxlength="5">頁<input type="submit" value="確定" name="B1">
</form>
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="646" cellspacing="1" cellpadding="0">
<tr>
<td width="10" class="p2"></td>
<td width="100" class="p2">作者</td>
<td width="336" class="p2">內容概要</td>
<td width="60" class="p2">天氣</td>
<td width="121" class="p2">日期</td>
<td width="70" class="p2">心情</td>
<td width="45" class="p2">點擊量</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

%>
<tr>
<%
if len(rs("diarymessage"))>20 then
diarymessage=cutstr(rs("diarymessage"),20)
else
diarymessage=rs("diarymessage")
end if
%>
<td width="10" class="p3"></td>
<td width="100" class="p3"><a
href="../areauserinfo.asp?user=<%=rs("user")%>"><%=rs("user")%></a></td>
<td width="336" class="p3"><a href="readjump.asp?id=<%=rs("id")%>"><%=diarymessage%></a></td>
<td width="60" class="p3"><%=rs("diaryseason")%> </td>
<td width="121" class="p3"><%=rs("diarydate")%>&nbsp;&nbsp;<%=rs("diarytime")%></td>
<td width="70" class="p3"><%=rs("mood")%> </td>
<td width="45" class="p3"><%=rs("clicknumber")%> </td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count
%>
<tr>
<td width="10" class="p3"> </td>
<td width="100" class="p3"></td>
<td width="336" class="p3"></td>
<td width="60" class="p3"> </td>
<td width="121" class="p3"> </td>
<td width="70" class="p3"> </td>
<td width="45" class="p3"> </td>
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