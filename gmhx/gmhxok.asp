<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
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
name=lcase(trim(request.form("name")))
pass=request.form("pass")
name1=lcase(trim(request.form("name1")))
for i=1 to len(name1)
usernamechr=mid(name1,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=120"
next
if name1="無" or name="無" or name="未定" or name1="未定" then Response.Redirect "../error.asp?id=130"
if left(name1,3)="%20" OR InStr(name1,"=")<>0 or InStr(name1,"`")<>0 or InStr(name1,"'")<>0 or InStr(name1," ")<>0 or InStr(name1," ")<>0 or InStr(name1,"'")<>0 or InStr(name1,chr(34))<>0 or InStr(name1,"\")<>0 or InStr(name1,",")<>0 or InStr(name1,"<")<>0 or InStr(name1,">")<>0 then Response.Redirect "../error.asp?id=120"
if chuser(name1) then Response.Redirect "../error.asp?id=120"
if instr(name,"or")<>0 or instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then Response.Redirect "../error.asp?id=120"
if instr(name1,"or")<>0 or instr(name1,"'")<>0 or instr(name1,"|")<>0 or instr(name1," ")<>0 then Response.Redirect "../error.asp?id=120"
'if Instr(name1,"朋友名字")>0 or Instr(name1,Application("sjjh_automanname"))>0 or Instr(name1,"測試測試")>0 or Instr(name1,"站長")>0 or Instr(name1,"網管")>0 or Instr(name1,"無")>0 or Instr(name1,"測試")>0 or Instr(name1,"媽")>0 or Instr(name1,"爸")>0 or Instr(name1,"大家")>0 or Instr(name1,"操")>0 then Response.Redirect "../error.asp?id=130"
if name="" or name1="" or pass="" then Response.Redirect "../error.asp?id=56"
'if Instr(LCase(Application("sjjh_useronlinename"))," "&LCase(name)&" ")<>0 then Response.Redirect "../error.asp?id=61"
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'校驗用戶
sql="SELECT id FROM 用戶 WHERE 姓名='" & name1 & "'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  Response.Redirect "../error.asp?id=62"
sql="SELECT 銀兩,身份,門派,grade,會員等級,性別,配偶 FROM 用戶 WHERE 姓名='" & name & "' and 密碼='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then Response.Redirect "../error.asp?id=63"
if rs("銀兩")<1000000 then Response.Redirect "../error.asp?id=457"
if rs("身份")="掌門" or rs("門派")="官府" or rs("grade")>=6 then Response.Redirect "../error.asp?id=64"
if rs("會員等級")>0 then Response.Redirect "../error.asp?id=65"
sex=rs("性別")
bl=rs("配偶")
sql="update 用戶 set 姓名='" & name1 & "',銀兩=銀兩-1000000 where 姓名='" & name & "'"
conn.Execute sql
sql="update 用戶 set 師傅='" & name1 & "' where 師傅='" & name & "'"
conn.Execute sql
if bl<>"無" and bl<>"" then
sql="update 用戶 set 配偶='" & name1 & "' where 姓名='" & bl & "'"
conn.Execute sql
if sex="男" then
conn.execute "update 合體技 set 姓名男='"& name1 &"' where 姓名男='" & name & "'"
else
conn.execute "update 合體技 set 姓名女='"& name1 &"' where 姓名女='" & name & "'"
end if
end if
conn.execute "update 貸款 set 貸款人='"& name1 &"' where 貸款人='" & name & "'"
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& name &"','"& name1 &"','操作','改換名字！')"
set rs=nothing
conn.close
set conn=nothing
'BBS論談
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
bbs="DBQ="+server.mappath("../bbs/bbs.asa")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open bbs
'校驗用戶
sql="SELECT * FROM book WHERE username='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update book set username='"&name1&"' where username='"&name&"'"
end if
set rs=nothing
conn.close
set conn=nothing
'記錄
'BBS論談
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
gupiao="DBQ="+server.mappath("../gupiao/gp.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open gupiao
'校驗用戶
sql="SELECT * FROM 大戶 WHERE 帳號='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update 大戶 set 帳號='"&name1&"' where 帳號='"&name&"'"
end if
sql="SELECT * FROM 客戶 WHERE 帳號='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update 客戶 set 帳號='"&name1&"' where 帳號='"&name&"'"
end if
'記錄
set rs=nothing
conn.close
set conn=nothing
'記錄
Response.Redirect "../ok.asp?id=111"
%>

 