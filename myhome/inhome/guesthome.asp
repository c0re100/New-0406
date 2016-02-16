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


<!--#include file="data2.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("name"),"名字")
%>
<%
hostname=request.querystring("name")
guestname=session("user")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖–訪問-客廳</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" text="#E18C59">

<%
rs.open "select * from userinfo where user='"&hostname&"' and homeopen='1'",conn,1,1

if rs.bof then
%>
<script language="Vbscript">
msgbox"這間房門緊閉著，似乎一直沒有人住過",0,"提示"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if
%>

<div align="center">
<center>
<table border="0" width="776" cellspacing="0" cellpadding="0">
<tr>
<td width="100%">&nbsp;</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="336"><font size="2">□-您當前的位置-<a
href="city.asp"><b><font color="#FFFF00">皇城中心</font></b></a>-<%=hostname%>的家</font></td>
<td class="p1" width="440"><font size="2">□-當前時間：<%=date%><%=time%></font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p2" width="496" colspan="4" height="21"><font size="2">你已經在<%=hostname%>家裡了，主人現在好像沒有在家</font></td>
</tr>
<tr>
<td class="p2" width="102" height="15"><font size="2">&nbsp;主人的信息：</font></td>
<td class="p2" width="64" height="15"></td>
<td class="p2" width="68" height="15"><font size="2">&nbsp;照片：</font></td>
<td class="p2" width="248" height="18"><font size="2">&nbsp;
主人給訪者的話：</font></td>
<td class="p2" width="262" height="18" colspan="3"><font size="2">&nbsp;主人開放的區域</font></td>
</tr>
<tr>
<td class="p3" width="102" height="16"><font size="2">&nbsp;姓名：</font></td>
<td class="p3" width="64" height="16"><font size="2"><%=rs("user")%></font></td>
<td class="p3" width="68" height="55" rowspan="4" align="center"><font
size="2"><img
src="../register/skin/<%if rs("skin")<>"" then
skin=rs("skin")
else
skin="01.gif"
end if
response.write skin%>"
alt="<%=rs("user")%>"></font></td>
<td class="p3" width="250" height="151" rowspan="8" valign="top"><font
size="2"><%=rs("guestmemo")%></font></td>
<td class="p3" width="261" height="139" valign="top" colspan="3"
rowspan="8">
<%
rs.close
rs.open "select * from openlist where user='"&hostname&"' and open='是'",conn,1,1
%>
<table border="0" width="100%" cellspacing="1" cellpadding="0">
<tr>
<td class="p3" width="37%"><a
href="../guestbook/send.asp?name=<%=hostname%>"><font size="2"
color="#FFFF00">給主人留言</font></a></td>
<td class="p3" width="63%"></td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td class="p3" width="37%"><a
href="../homejump.asp?fun=<%=rs("function")%>&amp;hostname=<%=hostname%>"><font
size="2"><%=rs("function")%></font></a></td>
<td class="p3" width="63%"></td>
</tr>
<%
rs.movenext
loop%>
</table>
</td>
</tr>
<tr>
<td class="p3" width="102" height="4"><font size="2">&nbsp;年齡：</font></td>
<%
rs.close
rs.open "select * from userinfo where user='"&hostname&"'",conn,1,3

%>
</tr>
<tr>
<td class="p3" width="102" height="19"><font size="2">&nbsp;性別：</font></td>
<td class="p3" width="64" height="19"><font size="2"><%=rs("sex")%> </font></td>
</tr>
<tr>
<td class="p3" width="102" height="5"><font size="2">&nbsp;來訪人次：</font></td>
<%
guestcount=rs("guestcount")+1
rs.update "guestcount",guestcount
%>
<td class="p3" width="132" height="5" colspan="2"><font size="2"><%=guestcount%> </font></td>
</tr>
</table>
</center>
</div>
<%rs.close%>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class="p1" width="776"><font size="2">&nbsp;</font></td>
</tr>
</table>
</center>
</div>

</body>
<%conn.close%>

</html>