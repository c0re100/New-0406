<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from 用戶 where id="&info(9),conn
if rs.bof or rs.eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起你沒有登陸!');location.href = '../exit.asp';}</script>"
	Response.End
mycorp=rs("門派")
end if%>
<%if request("action")<>"add" then%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-cn">
<style type="text/css">
<!--
table {  font-size: 9pt}
td {  font-size: 9pt}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>大俠加入</title></head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" topmargin="0" leftmargin="0" background="../ljimage/11.jpg">
<div align="center"><center><form method="POST" action="add.asp"><table border="1" width="600" height="310" cellspacing="0" bgcolor="#DCE2FF" bordercolor="#988CD0">

<tr><td width="594" colspan="2" bgcolor="#FFFFFF">
<p align="center"><font face="新細明體" size="3"><b>請大俠詳細填寫你的資料</b></font>
</td></tr>
<tr><td width="594" colspan="2">

<div align="center"><center><table border="1" width="100%" bordercolorlight="#000000" cellspacing="0" cellpadding="4" bordercolordark="#FFFFFF" bgcolor="#F8EBD1">
<tr>
<td width="13%" align="center"><a href=main.asp?action=search>⊙-大俠列表</a></td>
<td width="13%" align="center">⊙大俠加入</td>
<td width="13%" align="center"><a href=userre.asp>⊙-修改資料</a></td>
<td width="13%" align="center"><a href="sf.asp?Keys=Login">⊙-超級管理</a></td>
<td width="13%" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高級搜索</a></td>
<td width="35%" align="center">
<center>現有<font color="#FF6600">    <!--#include file="zongshu.asp" -->
</font>位大俠加入</center>
</td>
</tr></table></center></div>

</td></tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>姓名：</td>
<td width="479" bgcolor="#DCE2FF">
<input name=action type=hidden value=add>
<input type="text" name="name" value="<%=uname%>" size="15" maxlength="12">
</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>門派：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="mode" value="<%=rs("門派")%>" size="15" maxlength="50">你在夢幻的門派</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>密碼：</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="password" size="15" maxlength="30">防止黑客，請不要和夢幻的登陸密碼相同</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>確認密碼：</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="repassword" size="15" maxlength="30">同上</td>
</tr>

<tr>
<td width="111" align="right">性別：</td>
<td width="479" bgcolor="#DCE2FF"><select name="sex" size="1" class="p9">
<option value="男">性別→男 </option>
<option value="女">性別→女 </option>
</select></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>信箱：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="mail" size="30" maxlength="50" value="<%=rs("信箱")%>"></td>
</tr>

<tr>
<td width="111" align="right">主頁：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="url" size="30" maxlength="50" value="http://"></td>
</tr>

<tr>
<td width="111" align="right">您的網絡尋呼機：</td>
<td width="479" bgcolor="#DCE2FF"><select name="icq_img" class="p9" size="1">
<option value="oicq.gif" selected>OICQ</option>
<option value="icq.gif">ICQ</option>
</select>
<input type="text" name="icqnum" size="10" maxlength="10" value="<%=rs("oicq")%>">
<a href="http://www.icq.com">ICQ下載(英文)</a>，
<a href="http://www.oicq.com">OICQ下載(中文)</a></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>您的生日：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="years" size="4" maxlength="4" value="19" class="p9">年
<input type="text" name="mons" size="2" maxlength="2">月
<input type="text" name="days" size="2" maxlength="2">日</td>
</tr>

<tr>
<td width="111" align="right">聯繫電話：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="likes" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">手機或呼機：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="gsm" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">郵政編碼：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="pc" size="16" maxlength="6"></td>
</tr>

<tr>
<td width="111" align="right">單位名稱：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="danwei" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">詳細地址：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="address" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">照片上傳：</td>
<td width="479" bgcolor="#ddE2FF"><font color=aa0000>請先加入後，再到[<a href=userre.asp>修改資料</a>]處進行操作<vinput type="file" name="photo" class="p9"> </font> </td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>近況簡介：</td>
<td width="479" bgcolor="#DCE2FF"><textarea name="doc" rows="3" cols="39"><%=rs("簽名")%></textarea> </td>
</tr>

<tr>
<td width="594" colspan="2">
<center><input type="submit" value="加入" name="B1">
&nbsp;&nbsp;&nbsp;
<input type="reset" value="清除" name="B2"></center></td>
</tr>


</table></center></form></div>

</body>
</html>
<%Else
Dim name,repassword,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,conn,sql,rs,DBPath,strSQL
Dim lists,xingzuo

name     = Server.HtmlEncode(Request.Form("name"))      '姓名
password = Server.HtmlEncode(Request.Form("password"))  '密碼
repassword = Server.HtmlEncode(Request.Form("repassword"))  '確認密碼
sex      = Server.HtmlEncode(Request.Form("sex"))       '性別
mail     = Server.HtmlEncode(Request.Form("mail"))      'E-Mail
URL      = Server.HtmlEncode(Request.Form("URL"))       '網站鏈接
icq_img  = Server.HtmlEncode(Request.Form("icq_img"))   '傳呼機圖片
icqnum   = Server.HtmlEncode(Request.Form("icqnum"))    '傳呼機號碼
AGE      = Server.HtmlEncode(Request.Form("age"))       '年齡
years    = Server.HtmlEncode(Request.Form("years"))     '生日的年
mons     = Server.HtmlEncode(Request.Form("mons"))      '生日的月
days     = Server.HtmlEncode(Request.Form("days"))      '生日的日
likes    = Server.HtmlEncode(Request.Form("likes"))     '電話
gsm      = Server.HtmlEncode(Request.Form("gsm"))       '手機
pc  	 = Server.HtmlEncode(Request.Form("pc"))   	'郵政編碼
danwei   = Server.HtmlEncode(Request.Form("danwei"))    '單位
address  = Server.HtmlEncode(Request.Form("address"))   '居住地
photo    = Server.HtmlEncode(Request.Form("photo"))     '相片
doc      = Server.HtmlEncode(Request.Form("doc"))       '留言
mode     = Server.HtmlEncode(Request.Form("mode"))      '班級
ip       = Request.ServerVariables("REMOTE_ADDR")
if repassword<>password  then _
response.write "密碼錯誤。<a href = javascript:history.back()>返回重輸</a>":response.end

'使用ServerVariables讀取用戶的IP地址，將用戶IP放如IP變量中
counter  = "0"
n=Year(date())                                           '取得函數Year年的年份放入n的變量中
y=Month(date())                                          '取的函數Month月的月份放入y的變量中
r=Day(date())                                            '取的函數Day天的天數放入r的變量中
s=Hour(time())                                           '取得函數Hour小時的小時數放入s變量中
f=Minute(time())                                         '取的函數Minute分的分鐘數值放入f變量中
m=Second(time())                                         '取的函數Second秒的秒數放入m變量中
if len(y)=1 then y="0" & y                               '如果y的數值等於一位數的話，則y等於0+Y變量，也就是說，把單數前都加一個0，比如01 or 02 or 03
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "年" & y & "月" & r & "日" & " " & weekdayname(weekday(date())) & " " & s & ":" & f & ":" & m
'sj 等於 X年X月X日空格 星期幾 幾點幾分幾秒
'WeekDay是提供一個1與7之間的數值，也就是星期幾
times    = sj
'#####################星座判斷#########################
'山羊座 （12/22-1/19） 水瓶座 （1/20-2/18） 雙魚座 （2/19-3/20） 白羊座 （3/21-4/19）
'金牛座 （4/20-5/20） 雙子座 （5/21-6/20） 巨蟹座 （6/21-7/22） 獅子座 （7/23-8/22）
'處女座 （8/23-9/22） 天枰座 （9/23-10/22） 天蠍座 （10/23-11/21） 射手座 （11/22-12/21）
if days="" or years="" or mons="" then
response.write "請輸入數值!<a href = javascript:history.back()>返回重輸</a>"
response.end
end if
if IsNumeric(mons)=false or IsNumeric(days)=false or IsNumeric(years)=false then
response.write "您輸入的不是數字。<a href = javascript:history.back()>返回重輸</a>"
response.end
end if
mons=cint(mons)
days=cint(days)
years=cint(years)

if mons=>13 or mons<=0 then _
response.write "月份錯誤，最多只有12個月，最少也有1月。<a href = javascript:history.back()>返回重輸</a>":response.end
if days=>32 or days<=0 then _
response.write "日期錯誤，請認真點填寫！！<a href = javascript:history.back()>返回重輸</a>":response.end
if years<=1899 or years>=year(now()) then _
response.write "年份是從1900年開始計算，當然，你也不能填寫2000年出生的。<a href = javascript:history.back()>返回重輸</a>":response.end

if (mons=12 and days>=22) or (mons=1 and days<=19) then
xingzuo= "山羊座"
elseif (mons=1 and days>=20) or (mons=2 and days<=18) then
xingzuo= "水瓶座"
elseif (mons=2 and days>=19) or (mons=3 and days<=20) then
xingzuo= "雙魚座"
elseif (mons=3 and days>=21) or (mons=4 and days<=19) then
xingzuo= "白羊座"
elseif (mons=4 and days>=20) or (mons=5 and days<=20) then
xingzuo= "金牛座"
elseif (mons=5 and days>=21) or (mons=6 and days<=20) then
xingzuo= "雙子座"
elseif (mons=6 and days>=21) or (mons=7 and days<=22) then
xingzuo= "巨蟹座"
elseif (mons=7 and days>=23) or (mons=8 and days<=22) then
xingzuo= "獅子座"
elseif (mons=8 and days>=23) or (mons=9 and days<=22) then
xingzuo= "處女座"
elseif (mons=9 and days>=23) or (mons=10 and days<=22) then
xingzuo= "天枰座"
elseif (mons=10 and days>=23) or (mons=11 and days<=21) then
xingzuo= "天蠍座"
elseif (mons=11 and days>=22) or (mons=12 and days<=21) then
xingzuo= "射手座"
end if

'response.write "當前年份:" & year(now())
'response.write "<br>平均計算您的歲數是：" & (year(now()) - years)

'###############################
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SQL="select * from list where name='"+name+"' order by ID ASC"
rs.open sql,conn,3,3
if rs.eof and rs.bof  then
elseif name = rs("name") Then
name = ""
Password=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('用戶名已存在，請重新登陸填寫！');" & Chr(13)
response.write "</script>" & Chr(13)
response.write "<center><font color=red><a href = javascript:history.back()>返回重輸</a></font></center>"
Response.End
End If
rs.close



'--------------------判斷是否寫入必填字段--------
if name="" or mail="" or password="" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('您沒有填寫名稱或密碼或信箱，請重新登陸填寫！');" & Chr(13)
response.write "window.document.location.href=javascript:history.back();"&Chr(13)
response.write "</script>" & Chr(13)
End IF

if icq_img<>"0" Then
If icqnum="" Or IsNumeric(icqnum)=False Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('OICQ或icq號碼輸入錯誤!請重新輸入！！！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if

end if

if icq_img = "0" then
If icqnum <> "" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('請選擇傳呼機類型！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if
if icqnum = "" then
icqnum = "-------"
icq_img= null
end if
End If

if doc="" then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('請輸入留言！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
end if

if likes = "" then
likes="--------"
end if

if url = "" then
url="--------"
end if

if address = "" then
address="--------"
end if

if pc = "" then
pc=0
end if

if danwei = "" then
danwei="--------"
end if

if gsm = "" then
gsm="-------"
end if

if photo = "" then
photo="-------"
end if

'--------------------正確的mail判斷函數----------
if IsValidEmail(mail)=False then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('錯誤的email輸入');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if


function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if

end function

'-------------寫入數據庫start-------------
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
'sql="select * from list order by ID ASC"
sql="list"
rs.open sql,conn,3,3
rs.addnew
photo = null
rs.fields("name")=name
rs.fields("password")=password
rs.fields("sex")=sex
rs.fields("mail")=mail
rs.fields("url")=url
rs.fields("icqnum")=icqnum
rs.fields("icq_img")=icq_img
rs.fields("xingzuo")=xingzuo
age=n-years
rs.fields("age")=age       '年齡需要根據生日修改，而且需要制作星座代碼
rs.fields("years")=years
rs.fields("mons")=mons
rs.fields("days")=days
rs.fields("likes")=likes
rs.fields("gsm")=gsm
rs.fields("pc")=pc
rs.fields("danwei")=danwei
rs.fields("address")=address
rs.fields("photo")=photo
rs.fields("doc")=doc
rs.fields("mode")=mode
rs.fields("ip")=ip
rs.fields("counter")=counter
rs.fields("times")=times
rs.update
rs.close
'-----------------------數據添加成功,頁面跳轉到main.asp-----------
Response.Redirect "main.asp?action=search"
End If
%>
