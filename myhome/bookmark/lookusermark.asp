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
if request.form("type")="" or request.form("type")="all" then
searchsql="select * from bookmark where open='是' and user='"&request.querystring("user")&"' order by number desc"
else
searchsql="select * from bookmark where open='是' and user='"&request.querystring("user")&"' and type='"&request.form("type")&"' order by number desc"
end if

rs.open searchsql,conn,1,1
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
<td width="475" height="26" align="center"><b><font size="2"
color="#E18C59">俠客寓所--書簽</font></b></td>
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
<td class="p1" width="630">□-您當前的位置-&nbsp;
俠客寓所---我的書簽 ---查看<%=request.querystring("user")%>的書簽&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a href="javascript:history.back(1)"><font size="3"
color="#FFCC00">返回</font></a></td>
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
<form method="get" action="lookusermark.asp">
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

<form method="get" action="lookusermark.asp">
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
<form method="get" action="lookusermark.asp">
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
<form method="get" action="lookusermark.asp">
<input type="hidden" name="page" value="<%=totalpage%>"><input
type="submit" value="尾頁"
style="font-family: Tahoma; font-size: 12px">
</form>

<%else%>
<input type="submit" value="尾頁"
style="font-family: Tahoma; font-size: 12px">
<%end if%> </td>
<td class="p2" width="424">
<form method="POST"
action="lookusermark.asp?user=<%=request.querystring("user")%>">
按分類查詢：<select size="1" name="type">
<option selected value="all">所有類型</option>
<option value="娛樂休閑">娛樂休閑</option>
<option value="商業經濟">商業經濟</option>
<option value="電腦網絡">電腦網絡</option>
<option value="醫療健康">醫療健康</option>
<option value="文學藝術">文學藝術</option>
<option value="教育就業">教育就業</option>
<option value="科學技術">科學技術</option>
<option value="體育健身">體育健身</option>
<option value="社會文化">社會文化</option>
<option value="旅遊交通">旅遊交通</option>
<option value="政法軍事">政法軍事</option>
<option value="生活服務">生活服務</option>
<option value="社會科學">社會科學</option>
<option value="參考資料">參考資料</option>
<option value="搜索引擎">搜索引擎</option>
<option value="新聞媒體">新聞媒體</option>
<option value="遊戲天地">遊戲天地</option>
<option value="音樂網站">音樂網站</option>
<option value="圖書網站">圖書網站</option>
<option value="軟件下載">軟件下載</option>
<option value="電子郵箱">電子郵箱</option>
<option value="免費資源">免費資源</option>
<option value="圖庫網站">圖庫網站</option>
<option value="硬件天地">硬件天地</option>
<option value="聊天交友">聊天交友</option>
<option value="網絡編程">網絡編程</option>
<option value="社區論壇">社區論壇</option>
<option value="網上教學">網上教學</option>
<option value="電子賀卡">電子賀卡</option>
</select> <input type="submit" value="查詢" name="B1"
style="font-family: Tahoma; font-size: 12px">
</form>
</td>
<td class="p2" width="102">當前頁：<%=dqpage%>
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
<td class="p2" width="191">網站名稱</td>
<td class="p2" width="105">分類</td>
<td class="p2" width="146">具體分類</td>
<td class="p2" width="78" align="center">點擊次數</td>
<td class="p2" width="58"></td>
<td class="p2" width="87">所屬市民</td>
<td class="p2" width="111">添加到我的書簽</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount
%>
<tr>
<td class="p3" width="10"></td>
<td class="p3" width="191"><a
href="goto.asp?id=<%=rs("id")%>&amp;type=<%=rs("type")%>&amp;user=<%=rs("user")%>&amp;url=<%=rs("url")%>"
target="_blank"><%=rs("name")%></a></td>
<td class="p3" width="105"><%=rs("type")%></td>
<td class="p3" width="146"><%=rs("emptytype")%></td>
<td class="p3" width="78" align="center"><%=rs("number")%></td>
<td class="p3" width="58"></td>
<td class="p3" width="87"><a
href="../../areauserinfo.asp?user=<%=rs("user")%>"><%=rs("user")%></a></td>
<td class="p3" width="111">
<form method="POST" action="addmymark.asp">
<input type="hidden" name="type" value="<%=rs("type")%>"><input
type="hidden" name="emptytype" value="<%=rs("emptytype")%>"><input
type="hidden" name="name" value="<%=rs("name")%>"><input
type="hidden" name="url" value="<%=rs("url")%>"><input
type="hidden" name="open" value="是"><input type="submit"
value="添加" name="B1"
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
<td class="p3" width="10">&nbsp;</td>
<td class="p3" width="191">&nbsp;</td>
<td class="p3" width="105">&nbsp;</td>
<td class="p3" width="146">&nbsp;&nbsp;</td>
<td class="p3" width="78" align="center">&nbsp;</td>
<td class="p3" width="58">&nbsp;&nbsp;</td>
<td class="p3" width="87"></td>
<td class="p3" width="111"></td>
</tr>
<%
next
end if
%>
<%rs.close
conn.close%>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="630" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="630">&nbsp;</td>
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
