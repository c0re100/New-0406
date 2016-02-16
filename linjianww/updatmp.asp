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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
zm=""
on error resume next
s=Request.QueryString("subject")
subject=request.form("subject")
comment=request.form("comment")
name=trim(request.form("name"))
jijing=trim(request.form("jijing"))
shihe=trim(request.form("shihe"))
if name<>"未定" then
set rs=server.createobject("adodb.recordset")
rs.open ("select * from 用戶 where 姓名='"&name&"'"),conn,0,1
if rs.eof then
name="未定"
Response.Redirect "../error.asp?id=441"
else
if rs("身份")="掌門" then
zm=rs("門派")
end if
end if
end if
if subject="" or (shihe<>"男" and shihe<>"女" and shihe<>"所有") then
Response.Redirect "../error.asp?id=442"
else

if name<>"未定" then
conn.execute("update 用戶 set 門派='"&subject&"',身份='掌門',grade=5 where 姓名='"&name&"'")
end if
sql="Update 門派 Set 門派='" & subject & "',適合='" & shihe & " ',簡介='" & comment & " ',掌門='" & name & " ',幫派基金="&jijing&" Where 門派='" & s & "'"
conn.execute(sql)
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','門派更新操作')"
conn.Execute(sql)
rs.close
set rs=nothing
Response.Redirect "../ok.asp?id=701"
end if
%> 