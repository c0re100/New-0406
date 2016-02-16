<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
selectaction=request.form("s1")
selecthouse=request.form("r1")
bigarea=request.form("area")
if selectaction="" or selecthouse="" then
response.redirect"warning.asp"
else

select case selecthouse
case "house1"
sploshtemp=800
case "house2"
sploshtemp=20000
case "house3"
sploshtemp=50000
case "house4"
sploshtemp=200000
case "house5"
sploshtemp=500000
case "house6"
sploshtemp=1000000
end select
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖-房屋交易</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" background="../ljimage/11.jpg">
<div align="center">
</div>
<div align="center">
<center>
<table border="0" width="468" bordercolor="#FFFFFF" cellspacing="1"
cellpadding="0" height="1">
<tr>
<td class="p3" width="415" height="30" valign="top">
<%
select case selectaction
case "buy"
set rs=conn.execute("select 銀兩 from 用戶 where 姓名='"&info(0)&"'")
set rs=conn.execute("select house,bigarea from userinfo where user='"&info(0)&"'")
if rs("house")<>"0" then
%>
<table>
<tr>
<td><font size="2">您隻能擁有一所房屋！！如果想要換房屋，請先賣掉以前的房屋！！您的房屋類型是<%=rs("house")%>所在區為：<%=rs("bigarea")%></font></td>
</table>
<%
else
set rs=conn.execute("select 銀兩 from 用戶 where 姓名='"&info(0)&"'")
if rs("銀兩")<sploshtemp then
%>
<table>
<tr>
<td><font size="2">您沒有足夠的的錢呀！！</font></td>
</table>
<%
else
sploshtemp=rs("銀兩")-sploshtemp
rs.close
conn.execute"update 用戶 set 銀兩='"&sploshtemp&"' where 姓名='"&info(0)&"'"
conn.execute"update userinfo set house='"&selecthouse&"',bigarea='"&bigarea&"',homeopen='1' where user='"&info(0)&"'"
'set ls=conn.execute("select * from trueinfo where user='"&info(0)&"'")
'if ls.eof then
'conn.execute"insert into trueinfo(user) values ('"&info(0)&"')"
'end if
conn.close
conn1.close
%>
<table>
<tr>
<td><font size="2">恭喜您！您已經成功的購買了房子！！ ：-）</font></td>
</table>
<%
end if
end if

case "sale"
set rs=conn.execute("select 銀兩 from 用戶 where 姓名='"&info(0)&"'")
set rs=conn1.execute("select house from userinfo where user='"&info(0)&"'")

if rs("house")<>selecthouse then
%>
<table>
<tr>
<td><font size="2">您沒有這樣的房屋可賣！</font></td>
</table>
<%
else
set rs=conn.execute("select 銀兩 from 用戶 where 姓名='"&info(0)&"'")
sploshtemp=sploshtemp/4
sploshtemp=rs("銀兩")+sploshtemp
rs.close
conn.execute"update 用戶 set 銀兩='"&sploshtemp&"' where 姓名='"&info(0)&"'"
conn.execute"update userinfo set house='0',bigarea=' 'where user='"&info(0)&"'"
conn.close
%>
<table>
<tr>
<td><font size="2">恭喜！您成功的賣出房子，拿到了原價1/4的銀子。這裡真黑：（</font></td>
</table>
<%
end if
end select
end if

%>
</td>
</tr>
<tr>
<td class="p2" width="415" height="74" valign="top"></td>
</tr>
</table>
</center>
</div>

</body>

</html>
