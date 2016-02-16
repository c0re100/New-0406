<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
guessnumber=request.form("guessnumber")
if guessnumber="" then response.redirect"../../error.asp?id=488"
if not isnumeric(guessnumber) then response.redirect"../../error.asp?id=487"
intguessnumber=guessnumber-int(guessnumber)
if gussnumber<0 or intguessnumber<>0 then response.redirect"../../error.asp?id=486"
if len(guessnumber)<>5 then response.redirect"../../error.asp?id=485"
guessdate=date
Set cn=Server.CreateObject("ADODB.Connection")
Set rst=Server.CreateObject("ADODB.RecordSet")
guess="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("guess.asa")
cn.Open guess
sql="select * from 購買者 where 號碼="&guessnumber&" and datediff('d','"&guessdate&"',購買日期)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then

sql="select * from 購買者 where 購買者='"&username&"' and datediff('d','"&guessdate&"',購買日期)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then
	sql="SELECT 購買者 FROM 購買者 WHERE 購買者='"&username&"'"
	set rst=cn.Execute(sql)
	If not(Rst.Bof OR Rst.Eof) Then
	sql="update 購買者 set 號碼="&guessnumber&",購買日期='"&guessdate&"' where 購買者='"&username&"'"
	set rst=cn.execute(sql)
	'累計金加100
	sql="update 中獎號碼 set 累計金額=累計金額+1000"
	set rst=cn.execute(sql)
	
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	sql="update 用戶 set 銀兩=銀兩-100 where 姓名='"&username&"'"
	set rs=conn.execute(sql)

set rs=nothing	
set conn=nothing
	response.redirect "buy.asp"
	else
	sql="INSERT into 購買者(購買者,號碼,購買日期)values('"&username&"',"&guessnumber&",'"&guessdate&"')"
	Set Rst=cn.Execute(sql)
	'累計金加100
	sql="update 中獎號碼 set 累計金額=累計金額+1000"
	set rst=cn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	
	sql1="update 用戶 set 銀兩=銀兩-100 where id="&info(9)
	set rs=conn.execute(sql1)
	set rs=nothing	
set conn=nothing
Response.Write "<script language=JavaScript>{alert('購買成功！');location.href = 'buy.asp';}</script>"
	Response.End

	end if
else
set rs=nothing	
set conn=nothing
	response.redirect "../../error.asp?id=490"
end if
else
set rs=nothing	
set conn=nothing
	response.redirect "../../error.asp?id=489"
end if
rs.close
set rs=nothing	
set conn=nothing
%>
