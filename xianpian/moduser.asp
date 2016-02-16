<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp"  -->
<%
Dim USERID
Dim strSQL
Dim List,action
Dim name,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei,repassword
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,xingzuo
USERID=Request.Querystring("USERID")
IF Session("username")="" or Session("password")="" or USERID="" Then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('用戶信息丟失，請重新登陸！');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if

strSQL="SELECT * FROM List WHERE"
strSQL = strSQL & " name LIKE '%" & Session("username") & "%'"
strSQL = strSQL & " and password LIKE '%" & Session("password") & "%'"
strSQL = strSQL & " and ID LIKE '%" & USERID & "%'"
strSQL =strSQL & " Order by ID DESC"

set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'數據錯誤處理
if list.eof and list.bof then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
Response.WRite "錯誤！沒有當前用戶！請重新登陸！"
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if

IF  Session("Username")<>list("name") and Session("password")<>list("password") and USERID<>list("id") Then             '如果用戶名稱等於NULL則證明沒有通過驗證
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('沒有登陸，請重新登陸！');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if
session("mynameup")=list("name")
Response.Write "OK!通過確認了..."                        '確認用戶已經通過確認！
Response.Write Session("Username")

action = Request.Querystring("action")
If action="xiu" Then
Response.Write "修改記錄"
name     = trim(Request.Form("name"))      '姓名
password = trim(Request.Form("password"))  '密碼
repassword = trim(Request.Form("repassword"))  '密碼re
sex      = trim(Request.Form("sex"))       '性別
mail     = trim(Request.Form("mail"))      'E-Mail
URL      = trim(Request.Form("URL"))       '網站鏈接
icq_img  = trim(Request.Form("icq_img"))   '傳呼機圖片
icqnum   = trim(Request.Form("icqnum"))    '傳呼機號碼
AGE      = trim(Request.Form("age"))       '年齡
years    = trim(Request.Form("years"))     '生日的年
mons     = trim(Request.Form("mons"))      '生日的月
days     = trim(Request.Form("days"))      '生日的日
likes    = trim(Request.Form("likes"))     '愛好
gsm      = trim(Request.Form("gsm"))       '手機
pc  	 = trim(Request.Form("pc"))   	   '郵政編碼
danwei   = trim(Request.Form("danwei"))    '單位
address  = trim(Request.Form("address"))   '居住地
photo    = trim(Request.Form("photo"))     '相片
doc      = trim(Request.Form("doc"))       '留言
mode     = trim(Request.Form("mode"))      '同學方式
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
set list=server.createobject("ADODB.RECORDSET")
list.open strSQL,conn,3,2
list("name")=name
list("password")=password
list("sex")=sex
list("mail")=mail
list("url")=url
list("icqnum")=icqnum
list("icq_img")=icq_img
age=n-years
list("age")=age       '年齡需要根據生日修改，而且需要制作星座代碼
list("years")=years
list("mons")=mons
list("days")=days
list("likes")=likes
list("gsm")=gsm
list("pc")=pc
list("danwei")=danwei
list("address")=address
'list("photo")=photo
list("doc")=doc
list("mode")=mode
list("ip")=ip
list("times")=times
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>個人信息--<%=List("name")%></title><style>
<!--
#tipBox {position: absolute;
width: 160px;
z-index: 100;
border: 1pt black solid;
font-family:新細明體;
font-size: 9pt;
background: ffffdf;
visibility: hidden}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
</head>

<body background="../ljimage/11.jpg">
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" cellpadding="3" bgcolor="#998CD2" bordercolor="#988CD8">
<tr>
<td width="100%">
<p align="center">您的信息</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" cellspacing="0" cellpadding="3" bordercolor="#988CD8" style="font-family: 新細明體; font-size: 9pt" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="66" align="center"><a href=main.asp?action=search>⊙-大俠列表</a></td>
<td width="66" align="center"><a href="add.asp">⊙-大俠加入</a></td>
<td width="66" align="center"><a href="userre.asp">⊙-修改資料</a></td>
<td width="66" align="center"><a href="sf.asp?Keys=Login">⊙-超級管理</a></td>
<td width="66" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高級搜索</a></td>
<td width="220" align="center"><font color="#996600">現有</font><font color="#FF6600">
<!--#include file="zongshu.asp" --> </font><font color="#996600">位大俠加入</font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
    <table border="1" width="600" cellspacing="0" style="font-family: 新細明體; font-size: 9pt" cellpadding="3" bordercolor="#988CD0" height="272">
      <tr> 
        <td width="590" bgcolor="#988CD0" colspan="2" height="12">這是您的修改的結果</td>
      </tr>
      <!-- <tr>
<td width="590" colspan="2" bgcolor="#EEF1F7" height="12">photo...</td>
</tr> -->
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">姓名：</td>
        <td width="433" height="12"><%=list("name")%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">門派：</td>
        <td width="433" height="12"><%=mode%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">性別：</td>
        <td width="433" height="12"><%=list("sex")%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">信箱：</td>
        <td width="433" height="12"><%=list("mail")%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">生日：</td>
        <td width="433" height="12"><%=years%>年<%=mons%>月<%=days%>日</td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">聯繫電話：</td>
        <td width="433" height="12"><%=likes%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="11">郵編/地址：</td>
        <td width="433" height="11"><%=address%></td>
      </tr>
      <tr> 
        <td width="149" align="right" bgcolor="#CCCCFF" height="12">簡短留言：</td>
        <td width="433" height="12"><%=doc%></td>
      </tr>
      <tr> 
        <td width="582" colspan="2" bgcolor="#CCCCFF" height="25"> 
          <p align="center">
            <input class="p9" type="submit" value="  返回交友列表  ">
          </p>
        </td>
      </tr>
    </table>
</center>
</div>
</body>

</html>
<%
list.close
response.write "<script language='javascript'>" & chr(13)
response.write "alert('已經成功修改用戶資料，請按確定返回！');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style>
<!--
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<title>修改資料--<%=List("name")%></title>
</head>

<body>
<form method="POST" action="moduser.asp?USERID=<%=USERID%>&action=xiu">
<div align="center">
<center>
<table border="0" width="600" bgcolor="#988CD0" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
<p align="center"><b>修 改 資 料</b></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" bordercolor="#988CD0" cellspacing="0" cellpadding="3" style="font-family: 新細明體; font-size: 9pt">
<tr>
<td bgcolor="#F8EBD1" align="center"><a href="main.asp?action=search">⊙-大俠列表</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="add.asp">⊙-大俠加入</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="userre.asp">⊙-修改資料</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="sf.asp?Keys=Login">⊙-超級管理</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高級搜索</a></td>
<td  align="center"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='相片上傳'>⊙-相片上傳</a></td>
</tr>
</table>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<a href style='cursor:hand;'onClick="window.open('upload.asp','_top','scrollbars=yes,resizable=yes,width=600,height=400')" title='上傳相片'>還沒有上傳相片</a>
<%else %> <font size=2><a href style='cursor:hand;'onClick="javascript:location.reload();" title='刷新'>您的靚照</a><img src="showimg.asp?id=<%=session("Username")%>"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='相片重新上傳'>我要修改</a>
<%End If%>

</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#DFE1FF" style="font-family: 新細明體; font-size: 9pt" cellspacing="0" cellpadding="3" bordercolor="#988CD0">
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">姓名：</td>
<td width="424"><input readonly maxLength="30" name="name" size="30" value="<%=list("name")%>"></td>
<%session("myname")=list("name")%>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">門派：</td>
<td width="424"><input maxLength="30" name="mode" size="30" value="<%=list("mode")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">修改密碼：</td>
<td width="424"><input maxLength="30" name="password" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">確認修改密碼：</td>
<td width="424"><input maxLength="30" name="repassword" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">性別：</td>
<td width="424"><select class="p9" name="sex" size="1">
<% IF list("sex")="男" Then%>
<option value="">請選擇性別</option>
<option selected value="男">男</option>
<option value="女">女</option>
<%ElseIF list("sex")="女" Then%>
<option value="">請選擇性別</option>
<option selected value="女">女</option>
<option value="男">男</option>
<%Else%>
<option selected value="">請選擇性別</option>
<option value="男">男</option>
<option value="女">女</option>
<%End IF%>
</select></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">信箱：</td>
<td width="424"><input maxLength="50" name="mail" size="30" value="<%=list("mail")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">主頁：</td>
<td width="424"><input maxLength="50" name="url" size="30" value="<%=list("URL")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">網絡尋呼機：</td>
<td width="424"><select class="p9" name="icq_img" size="1">
<% IF list("icq_img") = "icq.gif" Then%>
<option value="0">沒有網絡尋呼機</option>
<option selected value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%ElseIf list("icq_img") = "oicq.gif" Then%>
<option value="0">沒有網絡尋呼機</option>
<option value="icq.gif">ICQ</option>
<option selected value="oicq.gif">OICQ</option>
<%Else%>
<option selected value="0">沒有網絡尋呼機</option>
<option value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%End if%>
</select>號碼<input maxLength="10" name="icqnum" size="10" value="<%=list("icqnum")%>"><a href="http://www.icq.com">ICQ下載</a>
<a href="http://www.tencent.com">OICQ下載</a></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">您的生日：</td>
<td width="424"><input maxLength="4" name="years" size="4" value="<%=list("years")%>">年
<input maxLength="2" name="mons" size="2" value="<%=list("mons")%>">月
<input maxLength="2" name="days" size="2" value="<%=list("days")%>">日</td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">聯繫電話：</td>
<td width="424"><input maxLength="50" name="likes" size="30" value="<%=list("likes")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">手機或呼機：</td>
<td width="424"><input maxLength="50" name="gsm" size="30" value="<%=list("gsm")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">郵政編碼：</td>
<td width="424"><input maxLength="6" name="pc" size="30" value="<%=list("pc")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">單位名稱：</td>
<td width="424"><input maxLength="50" name="danwei" size="30" value="<%=list("danwei")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">詳細地址：</td>
<td width="424"><input maxLength="50" name="address" size="30" value="<%=list("address")%>"></td>
</tr>
<!--    <tr>
<td width="160" bgcolor="#CCCCFF" align="right">相片上傳：</td>
<td width="424"><input readonly name="photo" value="<%=list("photo")%>" type="file" disabled >&nbsp;
輸入Delete則刪除您的相片</td>
</tr> -->
<tr>
<td width="160" bgcolor="#CCCCFF" align="right" valign="top">簡短留言：</td>
<td width="424"><textarea cols="39" name="doc" rows="3" wrap="hard" ><%=list("doc")%></textarea></td>
</tr>
<tr>
<td width="584" bgcolor="#CCCCFF" colspan="2">
<p align="center">
<input class="p9" type="submit" value="  修改  ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  清除  "></td>
</tr>
</table>
</center>
</div>

</forM>

</body>

</html>
