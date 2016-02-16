<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,體力加,體力 from 用戶 where id="&info(9),conn
tlj=(rs("等級")*1500+5260)+rs("體力加")+1000
sm=rs("體力")
%>

<head>
<title>地獄</title>
</head>

<body bgcolor="#996699" text="#FFFFFF" oncontextmenu=self.event.returnValue=false LEFTMARGIN="0" TOPMARGIN="0">
<table width="493" border="0" cellspacing="0" cellpadding="0" align="CENTER" height="267"> 
<tr> <td height="267" width="508"> <table width="458" border="0" cellspacing="0" cellpadding="0" height="318"> 
<tr> <td height="298" width="1" ></td><td  background="../images/haidian_0928_2.jpg" width="486" align="center" height="298"><br> 
<div id="Layer2" style="z-index: 2; left: 222px; width: 267; position: absolute; top: 316px; height: 80"> 
<table cellspacing="0" border="0" width="251" bordercolorlight="#cccc99" bordercolordark="#FFFFFF" height="15"> 
<% if sm>=tlj then
response.write "你目前的狀態很好啊，不需要再煉了"
response.end
else%> <form method=POST action=diyu2.asp  onSubmit="this.ok.disabled=true"> <tr> 
<td height="1" width="247"> <div align="left">你的生命力為<%=sm%>, <%
if sm>=10000 then response.write "不需要煉了！再煉就會降到10000了":p=1
if sm<10000 and sm>0 then response.write "實行一級煉火！":p=2
if sm<=0 and sm>-5000 then response.write "實行二級煉火":p=3
if sm<=-5000 then response.write "生命垂危，三味真火！":p=4
%> </div><tr> <td width="247"> <div align="left">治療方法： <input name=yilao readonly size=12 value="<%
if p=1 then response.write "回去吧,不歡迎你"
if p=2 then response.write "一級煉火"
if p=3 then response.write "二級煉火"
if p=4 then response.write "三味真火"
%>"> <input name=ok type=submit value=確定> &nbsp; </div></td></tr> </form><%
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if%> </table></div><br> </td><td height="298" width="49"> </div> </td></tr> 
</table></td></tr> </table>
</body>
<html>