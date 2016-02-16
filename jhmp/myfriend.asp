<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
%>
<html>
<head>
<title>江湖拉人記錄</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body bgcolor="#006699" background="../ff.jpg" text="#000000" vlink="#FFFF00" alink="#99FFFF"
leftmargin="0" topmargin="0">
<p align="center"> <FONT COLOR="#000033">感謝你朋友這些人是你拉到我們江湖的！<br>
  只有你拉的人存點夠100點纔可以在這裡顯示出來！<br>
  <br>
  簡介：拉人可以增加你自己的點數，當你所拉到我們江湖的人存點大於100時,<br>
  你每個月都可以從他身上扒皮一定點數（是計算機自動按5%計算，並不影響所拉人的泡點），在月底的<br>
  時候是最多的。如果這一個月過去了，你將扒不到皮的，只有等一下個月了，此值不累加！所以請大家多拉人吧！</FONT> <br>
   
<table width="621" border="0" cellpadding="2" cellspacing="2" height="13" align="center" bordercolor="#990000">
  <tr bgcolor="#FF0000"> 
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,姓名,性別,門派,身份,lasttime,mvalue,grade,等級,保留2 FROM 用戶 where allvalue>100 and 介紹人='"& info(0) &"' order by -mvalue",conn
%>
    <td width="26" height="10"> 
      <div align="center"><font color="#FFFFFF">ID</font></div></td>
    <td width="61" height="10"> 
      <div align="center"><font color="#FFFFFF">姓名</font></div></td>
    <td width="26" height="10"> 
      <div align="center"><font color="#FFFFFF">性別</font></div></td>
    <td width="52" height="10"> 
      <div align="center"><font color="#FFFFFF">門派</font></div></td>
    <td width="58" height="10"> 
      <div align="center"><font color="#FFFFFF">身份</font></div></td>
    <td width="113" height="10"> 
      <div align="center"><font color="#FFFFFF">最後登陸</font></div></td>
    <td width="48" height="10"> 
      <div align="center"><font color="#FFFFFF">月泡點</font></div></td>
    <td width="42" height="10"> 
      <div align="center"><font color="#FFFFFF">管理級</font></div></td>
    <td width="46" height="10"> 
      <div align="center"><font color="#FFFFFF">戰鬥級</font></div></td>
    <td width="36" height="10"> 
      <div align="center"><font color="#FFFFFF">扒皮點</font></div></td>
    <td width="45" height="10"> 
      <div align="center"><font color="#FFFFFF">扒皮</font></div></td>
  </tr>
  <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
  <tr> 
    <td width="26" height="30"> <div align="center"><font color="#000000"><%=rs("ID")%></font></div></td>
    <td width="61" height="30"> <div align="center"><font color="#0000FF"><%=rs("姓名")%></font></div></td>
    <td width="26" height="30"> <div align="center"><font color="#000000"><%=rs("性別")%></font></div></td>
    <td width="52" height="30"> <div align="center"><font color="#000000"><%=rs("門派")%></font></div></td>
    <td width="58" height="30"> <div align="center"><font color="#000000"><%=rs("身份")%></font></div></td>
    <td width="113" height="30"> <div align="center"><font color="#000000"><%=rs("lasttime")%></font></div></td>
    <td width="48" height="30"> <div align="center"><font color="#FF0000"><%=rs("mvalue")%></font></div></td>
    <td width="42" height="30"> <div align="center"><font color="#000000"><%=rs("grade")%></font></div></td>
    <td width="46" height="30"> <div align="center"><font color="#000000"><%=rs("等級")%></font></div></td>
    <td width="36" height="30"> <div align="center"><font color="#FF0000"><%=int(rs("mvalue")*0.05)%></font></div></td>
    <td width="45" height="30"> <div align="center"><font color="#000000"> 
        <%
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,sj)=0 and rs("保留2")<>"扒皮"&Month(date()) and rs("mvalue")>20 then%>
        <a href="babi.asp?id=<%=rs("id")%>"><font color="#0000FF">扒皮</font></a> 
        <%else%>
        不能操作 
        <%end if%>
        </font></div></td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p align="center"> <font color="#000000"><br> <font size="+3">拉人總數:</font></font><font size="+3"><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font></font> 
<div align="center"></div>
</body>
</html>