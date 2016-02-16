<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name1")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name2")
name2=trim(name2)
name2=server.HTMLEncode(name2)
kill=request.form("kill")
kill=trim(kill)
kill=server.HTMLEncode(kill)
mess=Request.form("mess")
mess=trim(mess)
money=Request.form("money")
if not isnumeric(money) then response.redirect"../error.asp?id=464"
money=int(money)
if money<1000000 or money>10000000000 then 
%>
<script language=vbscript>
MsgBox "雇傭金不能超過100億和小於100萬"
location.href = "shashoulist.asp"
</script>
<%response.end
end if
if len(mess)>10 then 
%>
<script language=vbscript>
MsgBox "要說的話不能超過10個字符"
location.href = "shashoulist.asp"
</script>
<%response.end
end if
mess=server.HTMLEncode(mess)
if name1="" or name2="" or kill="" then
	Response.Write "<script language=JavaScript>{alert('各項注意填寫盡量不要為空!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if name1=name2 or name1=kill or name2=kill or name1<>info(0) then
Response.Write "<script language=JavaScript>{alert('你的姓名、雇的殺手、被殺的人名字不能相同!!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'校驗用戶
sql="SELECT 門派 FROM 用戶 WHERE 姓名='" & name2 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('沒這個人啊？搞錯了吧！!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if rs("門派")<>"刺客" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你要雇傭的殺手隻能是刺客組織的人？！!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
rs.close
set rs=nothing

sql="SELECT 等級,門派,銀兩 FROM 用戶 WHERE 姓名='" & info(0) & "'"
Set Rs=conn.Execute(sql)
if rs("等級")<3 or rs("門派")="官府" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的等級太低了，要戰鬥等級大於3級和非官府的纔可以雇傭殺手？！!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if rs("銀兩")<money then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你有那麼多錢嗎？！!');location.href = 'gushashou.htm';}</script>"
response.end
end if
sql="update 用戶 set 銀兩=銀兩-"&money&" where id="&info(9)
conn.execute sql
rs.close
sql="SELECT id FROM 用戶 WHERE 姓名='" & kill & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('江湖中沒有你要殺的人啊？搞錯了吧！!');location.href = 'gushashou.htm';}</script>"
	response.end
end if

sz = "'" & name1 & "','" & kill & "','" & name2 & "','" & mess & "'," & money & ",now()"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
into_db = "INSERT INTO 殺手 (姓名1,被殺者,姓名2,說明,殺人傭金,時間) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from 殺手 where now()-時間>10"
conn.execute sql
msg="<font color=FFD7F4>小道消息：</font><font color=FFD7F4>"&info(0)&"</font>說：<font color=FFD7F4>"&mess&"</font>，於是要殺<font color=FFD7F4>"&kill&"</font>不惜花<font color=FFD7F4>"&money&"</font>兩銀子雇傭刺客組織的高手<font color=FFD7F4>"&name2&"</font>,殺了<font color=FFD7F4>"&kill&"</font>後<font color=FFD7F4>"&name2&"</font>就可以去領錢了，哈哈！"

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
sd(199)="<font color=FFD7F4>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

	Response.Write "<script language=JavaScript>{alert('登記成功！!');location.href = 'shashoulist.asp';}</script>"
	response.end%>