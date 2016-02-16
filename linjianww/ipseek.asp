<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
ipdizi=request.form("ipdizi")
if ipdizi="" then%>
<html>
<head>
<title>靈劍江湖</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor=#FFFFFF background="0064.jpg">
<center>
  <font color="#000000"><b><font size="+6" color="#FF0000"> 
  <%ip=Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'查最後ip的地區
sql="select * FROM 地址 WHERE ip1<="& num &" and ip2>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="亞洲"
city="未知"
else
country=rs("國家")
city=rs("城市")
end if
if country="" then country="香港"
if city="" then city="未知"
last="地區:"& country &" 城市:"& city
set rs=nothing	
set conn=nothing
%>
靈劍江湖ip查找程序</font></b></font><br>
<br>
感謝追捕作者提供的ip數據庫程序(11.1日前)<br>
注：該程序僅收入國內ip地址！<br>
您的ip地址為：<%=ip%> <%=last%>
<form action=ipseek.asp method=post>
請輸入ip地址並回車:
<input  size=20 name=ipdizi><input type=submit value='確認'>
</form>
</center>
</body>
</html>
<%else
ip=ipdizi
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'查最後ip的地區
sql="select top 1 國家,城市 FROM 地址 WHERE ip1<="& num &" and ip2>="&num
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
country="亞洲"
city="未知"
else
country=rs("國家")
city=rs("城市")
end if
if country="" then country="香港"
if city="" then city="未知"
last="地區:"& country &" 城市:"& city
set rs=nothing	
set conn=nothing%>
<html>
<head>
<title>靈劍江湖</title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../JHIMG/Bk_hc3w.gif">
您所查找的ip地址為：<%=ip%><br>鋻定為：<%=last%>
<br>
<br>
<a href="ipseek.asp">返回</a>
</html>
<%end if%> 