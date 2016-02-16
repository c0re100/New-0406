<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=trim(request("id"))
you=trim(request("you"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 身份,門派 FROM 用戶 WHERE id="&info(9),conn
if rs("身份")="掌門" and rs("門派")=id then
	rs.close
    rs.open "SELECT 身份,門派 FROM 用戶 WHERE 姓名='"&you&"'",conn
 if trim(rs("身份"))="掌門"  then
rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		response.end
	 end if	
 if trim(rs("門派"))<>id  then
rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		response.end
	 end if
	message="成功的把" & you & "逐出" & rs("門派") & "！"
	conn.execute "update 門派 set 弟子數=弟子數-1 where 門派='" & id & "'"
	conn.execute "update 用戶 set 門派='遊俠',身份='弟子',grade=1 where 姓名='" & you & "'"
else
	rs.close
set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		response.end
'message="你不是這個門派的掌門"
end if

Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【門派消息】["&info(0)&"]</font><font color=FFD7F4>成功的把[" & you & "]逐出[" & rs("門派") & "]！</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<title> 靈劍江湖 </title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body background="../images/8.jpg" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FF0000" size="+3">弟子管理</font><br>
<br>
江湖告示：<%=message%> </div>
<hr>
</body>
</html>