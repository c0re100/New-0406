<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>蝴蝶谷</title><meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body BGCOLOR="#996699" background="../ff.jpg" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"> 
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 體力 from 用戶 where id="&info(9) ,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
  <script language=vbscript>
MsgBox "對不起，不能進行操作！你還不是江湖中人，滾。"
window.close()
</script> <%
else
sm=rs("體力")
conn.close
set conn=nothing
set rs=nothing

%> <font color="#000000">歡迎<%=info(0)%>來蝴蝶谷治療（醫療事故不負責任） </font> </div><table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" BACKGROUND="../IMAGES/BJ31.JPG"> <tr> <td bgcolor="#CCE8D6" BACKGROUND="../ljimage/downbg.jpg"> 
<table> <form method=POST action='yilao2.asp' onsubmit="this.ok.disabled=true"> 
<tr> <td>胡醫仙說：你的生命力為<%=sm%>。 <%
if sm>=1000 then response.write "不需要治療！":p=1
if sm<1000 and sm>0 then response.write "建議多喫些補品！":p=2
if sm<0 and sm>-500 then response.write "你身體有恙，建議治療！":p=3
if sm<-500 then response.write "你傷得不輕，如不治療將危及生命！":p=4
%> <tr> <td>治療方法：<input name="yilao" readonly size="8"
value="<%
if p=1 then response.write "無"
if p=2 then response.write "喫些補品"
if p=3 then response.write "一般治療"
if p=4 then response.write "搶救生命"
%>"> <input name="ok" type="submit" value="確定"> <INPUT TYPE="button" VALUE="關閉" ONCLICK="javascript:window.close();" NAME="button"></td></tr> 
<tr> <td colspan="2" style="font-size:9pt"> <hr> 全部的治療的目標都是將生命調到1000。<br> 1、滋補主要是針對生命值在0--1000間的，每兩銀子加生命值1.5<br> 
2、一般治療是針對生命值在-500--0間的，每兩銀子加生命值1.0<br> 3、全力搶救是針對生命值在-500以下的，每兩銀子加生命值0.5。</td></tr> 
</form></table></td></tr> </table><%end if%> 
</body>
</html>
