<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"
Response.Buffer=true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER")
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
name=Request.form("name")
name=trim(name)
if info(0)<>name then
Response.Write "<script language=JavaScript>{alert('對不起，名字不符！！');parent.history.go(-1);}</script>"
Response.End 
end if
oldpass=trim(Request.form("oldpass"))
getpass=trim(request.form("getpass"))
repass=trim(request.form("repass"))
mail=trim(Request.Form ("mail"))
repass=trim(repass)
if instr(oldpass,"'")<>0 or instr(oldpass,"|")<>0 or instr(oldpass," ")<>0 or instr(oldpass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('對不起，名字密碼輸入有錯！！');parent.history.go(-1);}</script>"
response.end
end if
if instr(getpass,"'")<>0 or instr(getpass,"|")<>0 or instr(getpass," ")<>0 or instr(getpass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('對不起，密碼保護輸入錯誤！！');parent.history.go(-1);}</script>"
response.end
end if
if instr(repass,"'")<>0 or instr(repass,"|")<>0 or instr(repass," ")<>0 or instr(repass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('對不起，密碼保護重輸錯誤！！');parent.history.go(-1);}</script>"
response.end
end if
if instr(mail,"'")<>0 or instr(mail,"|")<>0 or instr(mail," ")<>0 or instr(mail,",")<>0 then
Response.Write "<script language=JavaScript>{alert('對不起，郵箱格式不對！！');parent.history.go(-1);}</script>"
response.end
end if
if not(instr(mail,".")<>0 and instr(mail,"@")<>0 and len(mail)>6) then 
Response.Write "<script language=JavaScript>{alert('對不起，郵箱格式不對！！');parent.history.go(-1);}</script>"
Response.End 
end if
if oldpass="" or getpass="" or repass="" or mail="" then
message="對不起！程序出錯<br>密碼，郵箱不能為空"
elseif getpass<>repass then
message="對不起！程序出錯<br>密碼與確認密碼必須一致"
else
temppass=StrReverse(left(oldpass&"godxtll,./",10))
templen=len(oldpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
oldpass=replace(mmpassword,"'","B")
temppass=StrReverse(left(getpass&"godxtll,./",10))
templen=len(getpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
getpass=replace(mmpassword,"'","B")
'校驗用戶
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.Open ("select 密取 from  用戶 where id="&info(9)),conn
if rs("密取")<>"" and rs("密取")<>"1" then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您己經申請了密碼保護，不可再次申請！！');parent.history.go(-1);}</script>"
Response.End 
end if
rs.Close
rs.Open ("SELECT * FROM 用戶 WHERE 姓名='" & name & "' and 密碼='" & oldpass& "'"),conn
If Rs.Bof OR Rs.Eof Then
message="對不起，你的密碼不對"
else
if  rs("狀態")="眠" then
message="對不起！程序出錯<br>你現在休息中，不能修改密碼！"
else
conn.Execute ("update 用戶 set 密取='" & getpass & "',信箱='"&mail&"' where 姓名='" & name & "'")
message="恭喜您成功地申請了密碼保護！"
end if
end if
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
'conn.Execute ("insert into v(v1,v2,v3,v4,v5)values('"&username&"','"&sj&"','"&ip&"','"&Request.Form ("oldpass")&"','"&Request.Form ("pass")&"')")
rs.Close
set rs=nothing
conn.Close
set conn=nothing
end if
%> 
<link rel="stylesheet" href="setup.css">
<body bgcolor=#990000>
<p align="center"><font color="yellow">◇ 修 改 密 碼 ◇</font></p>
<table border=1 bgcolor="#009999" align=center width=254 cellpadding="10" cellspacing="13" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td height="180"> 
<table width=254 height="100%">
<tr align="center"> 
<td> <%=message%> 
<p align=center><a href=ljsafe.asp>返 回</a> </p>
</td></tr>
</table>
</td>
</tr>
</table>
</body>