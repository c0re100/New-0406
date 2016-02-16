<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
userid=Request.Form("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 用戶 where id="&userid
rst.Open sqlstr,conn,3,3
if Request.Form("submit")="刪除" then
if info(0)<>"江湖總站" then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
name=rst("姓名")
sqlstr="delete * from 用戶 where id="&userid
conn.Execute sqlstr
sqlstr="delete * from 物品 where 擁有者="&userid
conn.Execute sqlstr
sqlstr="delete * from 交易市場 where 擁有者="&userid
conn.Execute sqlstr
sqlstr="delete * from 技能列表 where 擁有者='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from 賣面 where 擁有者='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from 使用技能 where 姓名='"&name&"'"
conn.Execute sqlstr
call isin(name)
Response.Write "成功刪除id="&userid&"的用戶"
elseif Request.Form("submit")="刪除他的物品" then
name=rst("姓名")
sqlstr="delete * from 物品 where 擁有者="&userid
conn.Execute sqlstr
sqlstr="delete * from 交易市場 where 擁有者="&userid
conn.Execute sqlstr
sqlstr="delete * from 技能列表 where 擁有者='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from 賣面 where 擁有者='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from 使用技能 where 姓名='"&name&"'"
conn.Execute sqlstr
call isin(name)
Response.Write "成功刪除id="&userid&"的用戶的物品"
else
for i=1 to rst.Fields.Count-1
Response.Write rst.Fields(i).Name&"("&rst.Fields(i).Type&")"&rst.Fields(i).Value&"---->"&Request.Form(i+1)&"<br>"
if rst.Fields(i).type=202 then
rst.Fields(i).Value=cstr(Request.Form(i+1))
elseif rst.Fields(i).type=3 then
rst.Fields(i).Value=clng(Request.Form(i+1))
elseif rst.Fields(i).type=6 then
rst.Fields(i).Value=Request.Form(i+1)
elseif rst.Fields(i).Type=135 then
rst.Fields(i).Value=cdate(Request.Form(i+1))
end if	
next
rst.Update
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
function isin(name)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
gupiao="DBQ="+server.mappath("../gupiao/gp.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open gupiao
'校驗用戶
sql="SELECT * FROM 大戶 WHERE 帳號='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "delete * from 大戶 where 帳號='"&name&"'"
end if
sql="SELECT * FROM 客戶 WHERE 帳號='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "delete * from  客戶  where 帳號='"&name&"'"
end if
'記錄
end function
%>
<body background="0064.jpg">
<a href="fine.asp">返回</a> 