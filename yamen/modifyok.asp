<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("你輸入的認證碼不正確，應該輸入:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if
name=Request.form("name")
name=trim(name)
'name=HTMLEncode(name)
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

if name="" then
message="帳號不能為空"
elseif oldpass="" or pass="" or repass="" then
message="密碼不能為空"
elseif pass<>repass then
message="密碼與確認密碼必須一致"
else
temppass=StrReverse(left(oldpass&"godxtll,./",10))
templen=len(oldpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
oldpass=replace(mmpassword,"'","B")
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
'ip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")

'校驗用戶
sql="SELECT 狀態 FROM 用戶 WHERE 姓名='" & name & "' and 密碼='" & oldpass& "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="對不起，你的密碼不對"
else
if  rs("狀態")="眠" then
message="你現在休息中，不能修改密碼！"
else
sql="update 用戶 set 密碼='" & pass & "' where 姓名='" & name & "'"
conn.Execute(sql)
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& name &"','無','操作','改換用戶密碼！')"
conn.close
message="恭喜您成功地修改了密碼！"
end if
end if
end if
%>
<head><LINK href="../css.css" rel=stylesheet>
</head>
<body background="../bg.gif">
<table border="1" align="center" width="360" cellpadding="10"
cellspacing="13" background="../images/12.jpg"> <tr bgcolor="#FFFFFF"> <td background="../ljimage/downbg.jpg"> 
<table width="100%"> <tr> <td> <p align="center" style="font-size:14;color:red"><%=message%></p><p align="center"><a href="#" onclick="window.close()">[ 
關 閉 窗 口 ]</a></p></td></tr> </table></td></tr> </table>

</body> 
<html><script language="JavaScript">                                                                  </script></html>