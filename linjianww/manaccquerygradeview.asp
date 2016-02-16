<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
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
querygrade=Request.QueryString("grade")
if querygrade="" then querygrade=1
if Not(IsNumeric(querygrade)) then querygrade=1
if querygrade<0 then page=0
if querygrade>10 then querygrade=10
querygrade=int(querygrade)
page=Request.QueryString("page")
if page="" then page=1
if Not(IsNumeric(page)) then page=1
if page<1 then page=1
page=int(page)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT 姓名,grade,times,allvalue,lasttime,lastip FROM 用戶 WHERE grade=" & querygrade
rs.open sql,conn,1,1
rs.PageSize=30
totalrec=rs.RecordCount
totalpage=rs.PageCount
if page>totalpage then page=totalpage
if totalrec>0 then
rs.AbsolutePage=page
p=1+(page-1)*rs.PageSize
Dim show()
i=0
j=1
Do while (Not rs.Eof) and (i<rs.PageSize)
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5)
show(j)=rs("姓名")
show(j+1)=rs("grade")
show(j+2)=rs("times")
show(j+3)=rs("allvalue")
show(j+4)=rs("lasttime")
show(j+5)=rs("lastip")
i=i+1
j=j+6
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing%><html>
<head>
<title>帳號管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">【帳號管理】</font></h1>
<font color="#FF0000">【按級別列出所有帳號】</font> <br>
<a href="manaccquerygrade.asp">返回</a></div>
<hr noshade size="1" color=009900>
<b> 注意事項 </b><br>
共有 <font color="#FF0000"><%=querygrade%></font> 級的帳號 <font color=red><%=totalrec%></font>
個。<%if totalrec>0 then%>
<hr noshade size="1" color=009900>
<div align=center>
<%for i=1 to totalpage
if page=i then
Response.Write " [" & i & "]"
else
Response.Write " <a href=manaccquerygradeview.asp?grade=" & querygrade & "&page=" & i & ">[" & i & "]</a>"
end if
next%></div>
<hr noshade size="1" color=009900>
<table border="1" cellspacing="0" cellpadding="6" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center" width="460">
<tr bgcolor="#3399FF">
<td><font color="#FFFFFF">序</font></td>
<td><font color="#FFFFFF">用戶名</font></td>
<td><font color="#FFFFFF">等級</font></td>
<td><font color="#FFFFFF">次數</font></td>
<td><font color="#FFFFFF">經驗值</font></td>
<td><font color="#FFFFFF">最後時間</font></td>
<td><font color="#FFFFFF">最後ＩＰ</font></td>
</tr>
<%for i=1 to UBound(show) step 6%>
<tr>
<td><%=(page-1)*30+(i+5)/6%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=show(i+3)%></td>
<td><%=show(i+4)%></td>
<td><%=show(i+5)%></td>
</tr>
<%next%>
</table><%end if%>
</body>
</html> 