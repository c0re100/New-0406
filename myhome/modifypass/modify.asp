<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

 
<!--#include file="../data.asp"-->
<%if info(0)<>"" then
name=info(0)
end if%>

<%Response.buffer=true
oldpass=Request.form("oldpass")
pass=request.form("pass")
repass=request.form("repass")
repass=trim(repass)
if instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then
response.end
end if
if instr(oldpass,"'")<>0 or instr(oldpass,"|")<>0 or instr(oldpass," ")<>0 then
response.end
end if
if instr(pass,"'")<>0 or instr(pass,"|")<>0 or instr(pass," ")<>0 then
response.end
end if
if instr(repass,"'")<>0 or instr(repass,"|")<>0 or instr(repass," ")<>0 then
response.end
end if


if oldpass="" or pass="" or repass="" then
message="密碼不能為空"

elseif pass<>repass then
message="密碼與確認密碼必須一致"
else
'ip=Request.ServerVariables("REMOTE_ADDR")
%>
<%

'校驗用戶
sql="SELECT * FROM 用戶 WHERE 姓名='" & name & "' and 密碼='" & oldpass& "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="對不起，你的密碼不對"
else
if  rs("狀態")="眠" then
message="你現在休息中，不能修改密碼！"
else
sql="update 用戶 set 密碼='" & pass & "' where 姓名='" & name & "'"
conn.Execute(sql)
message="恭喜您成功地修改了密碼！"
end if
end if
conn.close
end if
%>

<head>
<style>td{font-size:14}</style>
<title></title>
</head>

<body>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>


</td></tr>
</table>
</td></tr>
</table>
</body>