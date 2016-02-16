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
if instr(id,"官府")<>0 then
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
my=trim(request("my"))
if my<>info(0) then
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 門派,等級,性別 FROM 用戶 WHERE id="&info(9),conn
if rs.eof or rs.bof then	
message="你非江湖中人，不要在些搗亂！"
else
if rs("門派")="" or rs("門派")="無" or rs("門派")="遊俠"  then
	if rs("等級")<2 then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('等級不夠，需要2級以上纔可以加入！');window.close();</script>"
		response.end
	end if
	sex=rs("性別")
	rs.close
	rs.open "SELECT 適合 FROM 門派 WHERE 門派='" & id & "'",conn
	if rs.eof or rs.bof then
		message="江湖中沒有" & id & "這個門派"
	else
		if trim(sex)=trim(rs("適合")) or trim(rs("適合"))="所有" then
			message="你成功地加入了" & id & "這個門派"
			conn.execute "update 用戶 set 門派='" & id & "', 身份='弟子',grade=1 where id="&info(9)
			conn.execute "update 門派 set 弟子數=弟子數+1 where 門派='" & id & "'"
			info(5)=id
		else
			message="這個門派不適合你"
		end if
	end if
else
	message="要加入" & id & "，請你離開你現在的門派"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(199)="<font color=FFD7F4>【門派消息】["&info(0)&"]</font><font color=FFD7F4>成功地加入了[" & id & "]這個門派</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<title>靈劍江湖 </title>
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #0000FF; font-family: "新細明體"; font-size: 9pt; text-decoration: none }
a:hover      { color: #0000FF; font-family: "新細明體"; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" background="../linjianww/0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#FF0000" size="+3">加入江湖門派</font></p>
<p align="center"><font color="#FFFFFF"><b><font color="#000000">靈劍江湖告示：<%=message%>
</font></b></font><font color="#000000"></font></p>
<hr>

</body>

</html>
