<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
yilao=request.form("yilao")
if yilao<>"無" then
'校驗用戶
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 操作時間,體力,銀兩 FROM 用戶 WHERE id="& info(9) &" and 體力<1000",conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起系統限制，請等["&s&"分鐘]再操作！');}</script>"
	Response.End
end if	

If Rs.Bof OR Rs.Eof Then
mess="你不能進行治療"
else
sm=rs("體力")
yl=rs("銀兩")
select case yilao
case "喫些補品"
bl=1.5
cd=1000-sm
case "一般治療"
bl=1
cd=1000-sm
case "搶救生命"
bl=0.5
cd=1000-sm
end select
fy=int(cd/bl)
sm=1000
if yl<fy then
fy=yl
sm=yl*bl
end if
conn.execute "update 用戶 set 體力=體力+" & sm & ",銀兩=銀兩-" & fy & ",操作時間=now() where id="&info(9)
mess="經過胡醫仙的治療，你花了" & fy & "兩銀子增加生命到" & sm & "點"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
else
mess="你不用治療"
end if

set rs=nothing

set conn=nothing
%>

<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
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
<p align="center"><a href="yilao.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
