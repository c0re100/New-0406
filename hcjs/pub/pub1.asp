<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
time1=request.form("time")+0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if time1=0 then
mess="請選擇正確的時間！"
elseif  time1<0 then
mess="請選擇正確的時間！"
elseif  time1>24 then
mess="請選擇正確的時間！"
else
Jname=info(0)
Jpass=request.cookies("pass")
yin=abs((time1)*10)
shi=0.0416*time1
rs.open "select 銀兩 from 用戶 where id=" & info(9) & "  and 狀態<>'死' and 狀態<>'獄'",conn
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
if rs("銀兩")<yin then
mess=Jname & "，你的銀兩不夠！"
else
conn.execute "update 用戶 set 銀兩=銀兩-" & yin & ",體力=體力+"& yin &",內力=內力+"&yin&",登錄=now()+" & shi & ",狀態='眠' where id="&info(9)
mess=Jname & "，你已經成功的到客棧休息了，請在" &time1& "小時後使用該帳號！"
session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.htm"
</script>
<%
end if
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
<title>靈劍江湖客棧</title></head>

<body bgcolor="#000000" background="../../images/8.jpg" text="#000000" link="#0000FF">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="pub.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>