<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
myname=request.cookies("Jname")
myname=info(0)
rs.open "delete * from 月老 where 時間<now()-5",conn,3,3
rs.open "SELECT 姓名1,姓名2,說明,時間 FROM 月老 order by 時間 DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>月老</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style>td           { font-size: 14; color: #000000 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" bgcolor="#660000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<a href="zhenfen.asp">我要登記</a><font color="#0000FF"><br>

<a href="#" onClick="window.open('yuanou.asp','lihun','resizable=no,width=460,height=120,scrollbars=no,Status=no,toolbar=no,left=300,top=300,menubar=no,fullscreen=no')">江湖怨偶</a>
</font> <br>
<font color="#0000FF"><a href="Stunt.asp" title="離婚要損失生命100"> 合體特技</a> </font>
<table width="550" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
<tr bgcolor="#f7f7f7">
    <td align="left"><font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>頁][<font
color="#990000"><b><%=musers()%></b></font>人求婚] [<a
href="yuelao.asp?page=<%=page-1%>"><font color="#990000">上一頁</font></a>][第<%=page%>頁][<a
href="yuelao.asp?page=<%=page+1%>"><font color="#990000">下一頁</font></a>]</font></td>
</tr>
<tr>
<td style="font-size:21;color:#FF1493" align="center"><font size="2"><b><font size="+2">求婚信息</font></b></font></td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding="2" width="100%"
bgcolor="#000000" bordercolorlight="#EFEFEF">
<tr bgcolor="#DFEDFD">
<td><font size="2">求婚者</font></td>
<td><font size="2">求婚對像</font></td>
<td><font size="2">真情表白</font></td>
<td><font size="2">時間</font></td>
<td><font size="2">回復</font></td>
</tr>
<%
count=0
do while not rs.eof and count<rs.PageSize
%>
<tr bgcolor="#f7f7f7">
<td><font size="2"><%=rs("姓名1")%></font></td>
<td><font size="2"><%=rs("姓名2")%></font></td>
<td><font size="2"><%=rs("說明")%></font></td>
<td><font size="2"><%=rs("時間")%> </font>
<td><font size="2">
<%if rs("姓名2")=myname then %>
<a
href="jiefen.asp?name=<%=rs("姓名1")%>">同意</a>
<%end if%>
</font></td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</td>
</tr>
</table>
</body>

</html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 姓名1 from 月老")
musers=tmprs("姓名1")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%> 