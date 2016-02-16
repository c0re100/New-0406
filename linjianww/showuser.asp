<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
username=Request.Form("search")
cz=trim(Request.form("sjcz"))
name=request.querystring("username")
if name<>"" then
username=name
cz="姓名"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
Set rst=Server.CreateObject("ADODB.RecordSet")
if cz="ID" then
username=int(username)
sqlstr="SELECT * FROM 用戶 where "&cz&"="&username
rst.Open sqlstr,conn
else
 if cz="等級" then
sqlstr="SELECT * FROM 用戶 where "&cz&"="&username
rst.Open sqlstr,conn
else
sqlstr="SELECT * FROM 用戶 where "&cz&"='"&username&"'"
rst.Open sqlstr,conn
end if
end if
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的人我們找不到！請查看是否正確！');history.back();</script>"
else
Response.Write "<form method=post action=updateuser.asp><table><tr><td>id(readonly)</td><td><input style='border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid' type=text readonly name=id value="&rst("id")&"></td></tr>"
for i=1 to rst.Fields.Count-1
Response.Write "<tr><td>"&rst.Fields(i).Name&"("&rst.fields(i).Type&")</td><td><input style='border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid' type=text name="&rst.Fields(i).Name&" value='"&rst.Fields(i).value&"'></td></tr>"
next
Response.Write "<tr><td><input type=submit value=更新 name=submit></td><td><input type=submit value=刪除 name=submit><input type=submit value=刪除他的物品 name=submit><input type=reset value=重置></td></tr>"
Response.Write "</table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>

<html>
<body background="0064.jpg">
</html> 