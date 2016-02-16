<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
jiu=request.form("jiu")
Jname=info(0)
select case jiu
case "lao"
tili=50
yin=250
case "wu"
tili=60
yin=300
case "du"
tili=70
yin=350
case "mao"
tili=80
yin=400
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,武功 from 用戶 where id=" & info(9) & " and 狀態<>'死' and 狀態<>'獄' ",conn
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
if rs("銀兩")<yin then
mess=Jname & "，你的銀兩不夠！"
else
wugong=rs("武功")
jiuliang=int((wugong)/100)
if jiuliang>1 then
conn.execute "update 用戶 set 銀兩=銀兩-" & yin & ",體力=體力+"& tili &" where id="&info(9)
mess=Jname & "好酒量!你喝完一杯,不停的贊道:好酒,好酒!"
else	
shi=0.0416*1
conn.execute "update 用戶 set 銀兩=銀兩-" & yin & ",體力=體力+"& tili &",登錄=now()+" & shi & ",狀態='眠' where id="&info(9)
mess=Jname & "，由於您酒量不夠,趴在桌上就睡著了,幾個伙計將您抬到客棧去休息了，請在1小時後使用該帳號！"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.asp"
</script>
<%
end if
end if
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000">

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

