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
sql="select * from �ʶR�� where ���X="&guessnumber&" and datediff('d','"&guessdate&"',�ʶR���)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then

sql="select * from �ʶR�� where �ʶR��='"&username&"' and datediff('d','"&guessdate&"',�ʶR���)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then
	sql="SELECT �ʶR�� FROM �ʶR�� WHERE �ʶR��='"&username&"'"
	set rst=cn.Execute(sql)
	If not(Rst.Bof OR Rst.Eof) Then
	sql="update �ʶR�� set ���X="&guessnumber&",�ʶR���='"&guessdate&"' where �ʶR��='"&username&"'"
	set rst=cn.execute(sql)
	'�֭p���[100
	sql="update �������X set �֭p���B=�֭p���B+1000"
	set rst=cn.execute(sql)
	
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	sql="update �Τ� set �Ȩ�=�Ȩ�-100 where �m�W='"&username&"'"
	set rs=conn.execute(sql)

set rs=nothing	
set conn=nothing
	response.redirect "buy.asp"
	else
	sql="INSERT into �ʶR��(�ʶR��,���X,�ʶR���)values('"&username&"',"&guessnumber&",'"&guessdate&"')"
	Set Rst=cn.Execute(sql)
	'�֭p���[100
	sql="update �������X set �֭p���B=�֭p���B+1000"
	set rst=cn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	
	sql1="update �Τ� set �Ȩ�=�Ȩ�-100 where id="&info(9)
	set rs=conn.execute(sql1)
	set rs=nothing	
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ʶR���\�I');location.href = 'buy.asp';}</script>"
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
