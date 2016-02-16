<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
 
<%if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
nickname=info(0)
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
who=Trim(Request.QueryString("who"))
if who="" then who=nickname
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
chatroombgimage=ljjh_chatroombgimage
chatroombgcolor=ljjh_chatroombgcolor
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade,門派 FROM 用戶 where 姓名='" & info(0) & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖人物！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
else
pai=rs("門派")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>化骨綿掌</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="ccss.css" rel="stylesheet" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed"  background=bk.jpg>
<div align="center">化骨綿掌<br>
戰鬥等級大於50級<br>
    花費銀兩20000000兩 <br>
    攻擊對像等級小於20級<br>
<form method=POST action='juezhao2ok.asp'>
攻擊對像
<select name="to1">
<%
for i=0 to x
if show(i)<>info(0) then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
    <br>
    <br>
     <input type=submit value="施展絕技" name=submit>
</form>
<form method=POST action='juezhao2ok.asp'>
    <select name="to1" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<<%
if session("nowinroom") = 1 then
for gw=5 to 20
gw=gw+3
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=20 to 150
gw=gw+15
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=180 to 390
gw=gw+30
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=490 to 780
gw=gw+50
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=801 to 1100
gw=gw+49
gw1=gw&"級怪物"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
    </select>
    <br>
    <input type=submit value="施展絕技" name=submit2>
  </form>
  
</div>
</body>
</html>
<%end if%>