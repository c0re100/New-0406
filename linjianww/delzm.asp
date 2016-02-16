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
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
on error resume next
s=Request.QueryString("username")
rs=conn.execute("select * from 門派 where 門派='"&s&"'")
zm=rs("掌門")
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
rs.open ("select * from 用戶 where 姓名='"&zm&"'"),conn,0,1
if rs.bof or rs.eof then
Response.Redirect "../error.asp?id=446"
else
conn.execute("update 用戶 set 身份='弟子',grade=1 where 姓名='"&zm&"'")
end if

conn.execute("update 門派 set 掌門='未定' where 門派='"&s&"'")
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','掌門開除')"
conn.Execute(sql)
rs.close
set rs=nothing
Response.Redirect "../ok.asp?id=703"
%> 