<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhzb")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(Request.form("xlsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=request.querystring("action")
name=info(0)
if action<>"" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
rs.close
rs.open "select 數量,ID,功效,增加數值 from 修練物品 where 物品名='" & action & "' and 擁有者=" & info(9) & " and 數量>0",conn
if rs.eof or rs.bof then
mess="你沒有這種物品！"
else
if int(rs("數量"))<sl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的物品只有"&rs("數量")&"個，而你想服用"&sl&"個,你的物品不夠！');</script>"
	%>
	<script language="vbscript">
	location.href = "javascript:history.go(-1)"
	</script>
	<%
	response.end
end if
id=rs("ID")
gx=rs("功效")
jjsj=int(rs("增加數值"))*sl
conn.execute "update 修練物品 set 數量=數量-"& sl &" where id=" & id
conn.execute "delete * from 修練物品 where 數量<=0"
if gx="殺人數" then
	conn.execute "update 用戶 set "& gx &"="& jjsj &",操作時間=now() where id="&info(9)
else
	conn.execute "update 用戶 set "& gx &"="& gx &"+"& jjsj &" ,操作時間=now() where id="&info(9)
end if
mess="你喫用了藥物"& action & sl &"個，增加" & gx & jjsj & "點！"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>江湖提示</b></font></td>
</tr>
<tr align="center">
<td>
<table width="260">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
</td>
</tr>
</table>
</td>
</tr>
<tr align="center">
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回裝備一覽</a></td>
</tr>
</table>
</body>