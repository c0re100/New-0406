<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
Response.End
end if

username=info(0)
function amount()
dim rs
rs=conn.execute("Select sum(交易量) As samount from 股票")
amount=rs("samount")
set rs=nothing
if isnull(amount) then amount=0
end function
%>
<!--#include file="jhb.asp"--><%
sql= "select * from 股票"
set rss=conn.execute(sql)
tai=rss("狀態")

%>
<html>
<head>
<title>笑傲江湖股票交易所</title>
<style type="text/css"><!--td {  font-family: 新細明體; font-size: 9pt}body {  font-family: 新細明體; font-size: 9pt}select {  font-family: 新細明體; font-size: 9pt}A {text-decoration: none; font-family: "新細明體"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "新細明體"; font-size: 9pt} .big {  font-family: 新細明體; font-size: 10.8pt}
--></style>
</head>

<body bgcolor="#FFFEF4">
<div align="center"><center>

<table border="0" width="100%" height="41" cellpadding="0" cellspacing="0">
<tr>
<td width="10%" height="16" bgcolor="#FFFAE1"><img border="0" src="logo.jpg" width="100" height="33"></td>
<td width="423" height="16" bgcolor="#FF6699">
<p align="center"><b><font color="#FBE651" face="黑體" size="3">笑傲江湖股票交易所</font> </b>
</td>
<td width="10%" height="16" bgcolor="#FFFAE1"><img border="0" src="logo.jpg" width="100" height="33"></td>
</tr>
</table>
</center></div><div align="center"><center>

<table border="1" width="744" height="19" bordercolorlight="#FFFFFF" bordercolor="#000000"
cellspacing="0">
<tr>
<td width="744" height="19"><a href="jingji.asp">股票經紀人</a>---<a href="index.asp">股市行情更新</a>---<a href="../jh.asp">返回
</a>
<%if info(2)>9  then%>
---<a href="kai.asp">開盤</a>---<a href="feng.asp">收市</a>---<a href="chufa.asp">隨機事件</a>---<a href="jiu.asp">救市</a>---<a href="jiang.asp">給股市降溫</a>
<%end if%>
</td>
</tr>
</table>
</center></div>

<table border="0" width="100%" bgcolor="#FFF0B2" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFAE1">
<td width="100%" height="20">今日大盤顯示 -----股票開市以來成交總量<font
color="red"><%=amount()%></font>-----
目前股市狀態為<font
color="red"><%=tai%>盤</font><br>
股票以當天的開盤價格價格為基準，高於為漲，底於為跌，漲跌以交易量和隨即的運氣為標準。漲停和跌停的範圍都為50%</td>
</tr>
</table>

<table border="0" width="100%" bgcolor="#FBE651" cellspacing="1" cellpadding="0">
<tr>
<td bgcolor="#006633" height="20" align="center" width="141"><font color="#FFFFFF">股票代號</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">股票名稱</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">開盤價格</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">當前價格</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">漲/跌</font></td>
<td bgcolor="#006633" height="20" align="center" width="113"><font color="#FFFFFF">成交量</font></td>
</tr>
<%
sql= "select * from 股票"
set rs=conn.execute(sql)
DO UNTIL RS.EOF  %>
<%if CSng(rs("開盤價格"))<CSng(rs("當前價格")) then
color="#FF0000"
elseif CSng(rs("開盤價格"))>Csng(rs("當前價格")) then
color="#00FF00"
else
color="#FFFF00"
end if

if (rs("當前價格")-rs("開盤價格"))/rs("開盤價格")>=0.5 then
sstop=1
elseif (rs("當前價格")-rs("開盤價格"))/rs("開盤價格")<=-0.5 then
sstop=2
else
sstop=0
end if

%>
<tr bgcolor="#000000">
<%if sstop=1 then%>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%>(漲停)</a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("企業")%></a></span></font></td>
<%elseif sstop=2 then%>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%>(跌停)</a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("企業")%></a></span></font></td>
<%else %>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>">00<%=rs("sid")%></a></span></font> </td>
<td height="20" width="261"><span class="p"><a href="exchange.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("企業")%></a></span></font></td>
<%end if%>
<%
if rs("流通股票")<=1000000 then
randomize time()
addstocknum=int(rnd*1000000+1000000)
end if
%>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("開盤價格"),2)%> </span></font></td>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("當前價格"),2)%> </span></font></td>
<td height="20" width="148"><font color="<%=color%>"><span class="p"><%=formatnumber(rs("當前價格")-rs("開盤價格"),2)%> </span></font></td>
<td height="20" width="113"><font color="<%=color%>"><span class="p"><%=rs("交易量")%> </span></font></td>
</tr>
<%
rs.MoveNext
Loop
set rs=Nothing
%>
<%
sql="select * from 客戶 where 帳號='"&username&"'"
set rs=conn.execute(sql)
if not(rs.eof or rs.bof) then
uname=rs("帳號")
%>
<table border="0" width="100%" bgcolor="#FFF0B2" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFAE1">
<td width="100%" height="20">□- 下面是大俠<%=username%>您所擁有的股票，<font color="red">↓股票的買賣要慎重哦，要不白交1%的傭金了：）</font></td>
</tr>
</table>
<%
set rs=nothing
%>
<table border="0" width="100%" bgcolor="#FBE651" cellspacing="1" cellpadding="0">
<tr>
<td bgcolor="#006633" height="20" align="center" width="141"><font color="#FFFFFF">股票代號</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">股票名稱</font></td>
<td bgcolor="#006633" height="20" align="center" width="261"><font color="#FFFFFF">擁有數量</font></td>
<td bgcolor="#006633" height="20" align="center" width="148"><font color="#FFFFFF">平均價格</font></td>
</tr>
<%
sql="select * from 大戶 where 帳號='"&uname&"' and 持股數<>0"
set rs_u=conn.execute(sql)
DO UNTIL RS_U.EOF

%>
<tr bgcolor="#000000">
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs_u("sid")%>"><font color=Violet><%=rs_U("sid")%></a></span></font> </td>
<td height="20" width="141"><span class="p"><a href="exchange.asp?sid=<%=rs_u("sid")%>"><font color=Violet><%=rs_U("企業")%></a></span></font> </td>
<td height="20" width="148"><font color=Violet><span class="p"><%=rs_u("持股數")%> </span></font></td>
<%if rs_u("平均價格")=0 then%>
<td height="20" width="148"><font color=Violet><span class="p"><%=formatnumber(rs_u("買入價格"),3)%> </span></font></td>
<%else%>
<td height="20" width="148"><font color=Violet"><span class="p"><%=formatnumber(rs_u("平均價格"),3)%> </span></font></td>
<%end if%>
</tr>
<%
rs_u.MoveNext
Loop
end if
conn.close
set conn=nothing
%>

</table>
</body>
</html>