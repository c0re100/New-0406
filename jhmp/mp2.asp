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
rs.open "SELECT 門派,身份,存款 FROM 用戶 WHERE id="&info(9),conn
if rs.eof or rs.bof then
	message="你非江湖中人"
else
	if rs("門派")="" or rs("門派")="遊俠" or rs("門派")="無" then
		message="你並無門派"
	else
		if rs("身份")="掌門" then 
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('不想作了說一聲，是不是想死！！');window.close();</script>"
			response.end
		end if
if info(4)=0 then
	message="你離開了" & rs("門派") & "，損失內力1000,銀兩及存款均降到1000兩！"
	if rs("存款")>1000 then
		conn.execute "update 用戶 set 門派='遊俠',身份='弟子', 內力=內力-1000,grade=1,銀兩=1000,存款=1000 where id="&info(9)
	else
		conn.execute "update 用戶 set 門派='遊俠',身份='弟子', 內力=內力-1000,grade=1,銀兩=0 where id="&info(9)
	end if
else
	message="你離開了" & rs("門派") & "，因為你是江湖會員，所以原有內力，金錢不變！"	
	conn.execute "update 用戶 set 門派='遊俠',身份='弟子',grade=1 where id="&info(9)
end if
	conn.execute "update 門派 set 弟子數=弟子數-1 where 門派='" & id & "'"
	info(5)=""
end if

end if
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

<body bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"><font size="+3" color="#FF0000">離開門派</font><br>
<br>
<br>
靈劍江湖告示：<%=message%> </div>
</body>
</html>
