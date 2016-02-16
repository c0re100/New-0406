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
<!--#include file="data.asp"-->
<%
if Session("myname")="" then Response.Redirect "../error.asp?id=210"
date1=request.form("date")+0
time1=request.form("time")+0
if date1=0 and time1=0 then
mess="請選擇正確的時間！"
elseif date1<0 or time1<0 then
mess="請選擇正確的時間！"
elseif date1>3 or time1>24 then
mess="請選擇正確的時間！"
else
sql1="select house from userinfo where user='" & session("myname") & "'"
set rs1=conn1.execute(sql1)
housetype=rs1("house")
rs1.close
sql2="select sleepenergy from rest where housetype='"&housetype&"'"
set rs2=conn1.execute(sql2)
sleepenergy=rs2("sleepenergy")
rs2.close
Jname=session("myname")
yin=(date1*24+time1)*sleepenergy	
shi=date1+0.035*time1
sql="select 體力加,體力,內力加,內力 from 用戶 where 姓名='" & Jname & "'  and 狀態<>'死' and 狀態<>'獄'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
timax=rs("體力加")
tinow=rs("體力")
neimax=rs("內力加")
neinow=rs("內力")
nei=yin
tini=yin
if tinow+tini>timax then tini=timax-tinow
if neinow+nei>neimax then nei=neimax-neinow
sql="update 用戶 set 體力=體力+"& tini &",內力=內力+"&nei&",登錄=now()+" & shi & ",狀態='眠' where 姓名='" & Jname & "'"
conn.execute sql
mess=Jname & "，你已經鼾聲入雷了：），您即將增加體力" &tini&"內力"&nei&"請在" & date1 & "天" &time1& "小時後使用該帳號！"
session("myname")=""
Session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../index.htm"
</script>
<%
	
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000" background="../ljimage/11.jpg">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="sleep.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
