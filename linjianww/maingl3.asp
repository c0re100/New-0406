<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="update 用戶 set 會員費=會員費+500,金幣=金幣+500 where 會員等級=1 "
Set Rs=conn.Execute(sql)
sql="update 用戶 set 會員費=會員費+1000,金幣=金幣+1000 where 會員等級=2 "
Set Rs=conn.Execute(sql)
sql="update 用戶 set 會員費=會員費+1500,金幣=金幣+1500 where 會員等級=3 "
Set Rs=conn.Execute(sql)
sql="update 用戶 set 會員費=會員費+2000,金幣=金幣+2000 where 會員等級=4 "
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('現在已將所有會員加了會費！');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>

 