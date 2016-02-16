<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp" -->
<%
dim copyright
dim Aff,I
dim list,MaxPage,EditUserID,xiu
Aff = Chr(34)
Dim strSQL,Keys,UserID,Action
copyright = "<center><font style=" & Aff & "color: #000000; font-family: 新細明體; font-size: 9pt" & Aff & ">" & vbCrLf
copyright = copyright & "Copyright &copy; 2000-2003  " & vbCrLf
copyright = copyright & " All rights reserved<br>" & vbCrLf
copyright = copyright & "版權所有：<a href=" & Aff & "http://21ex.com" & Aff & " target=" & Aff & "_blank" & Aff & ">夢幻工作室</a>" & vbCrLf
copyright = copyright & "網站維護：<a href=" & Aff & "mailto:21ex@21cn.com" & Aff & "></a>"  & vbCrLf
copyright = copyright & "</font></center>" & vbCrLf
'程序開始，有三個判斷
'1、判斷超級用戶帳號是否為空
'2、判斷是否為超級用戶
'3、進行用戶管理
'4、注銷超級用戶身份
'首先查看表是否有管理員身份在內，如果沒有則要求注冊。
'如果表有管理員資料則和提交的表單的身份進行驗證
'如果驗證通過即可以進入管理界面
'如果不通過則輸出錯誤信息，本返回index頁面
'如果退出管理環境則注銷當前用戶。
'//showpage_admin_login             '//子程序用來顯示管理員登陸頁面
'//ShowRegAdminPage                 '//子程序用來顯示注冊管理員頁面
'//SQL_Server                       '//數據庫操作
'//DeleteOrEditOfAdmin              '//子程序用來顯示管理用戶頁面
'//chk_admin_password               '//檢測登陸用戶的密碼是否正確
'//Write_admin_data                 '//寫入數據庫超級用戶的信息
'//delete_user_list                 '//刪除用戶
'//CHK_admin_ISNULL                 '//檢查數據表是否為空字段
'------------------Program Start --------------------------------------
'目前需要兩個表示不同處理的關鍵字

Keys=Request("Keys")
Response.Write Keys
Select case Keys
case "Login"
CHK_admin_ISNULL                   '//檢測輸出什麼頁面
case "RegAdmin"
Write_admin_data                   '//寫admin數據
case "ChkAdmin"
CHK_chk_admin_password_OUTPage     '//檢測是否為sysop
case "DeleteID"                    '//刪除用戶操作字符
delete_user_list
case "EditUser"                    '//編輯用戶資料
IsAdminEditUser
case Else
%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Help</title>
</head>
<body background="../ljimage/11.jpg">
<div align="center"><center><table border="1" cellspacing="3" cellpadding="0" style="font-family: 新細明體; font-size: 10.5pt; color: #000000" bgcolor="#FF9900" bordercolorlight="#FFCC00" bordercolordark="#663300" bordercolor="#FF9900" width="218"><tr><td width="206" colspan="2">
<p align="center">【大俠錄參數說明】
</td></tr><tr><td width="56">Login-</td><td width="144">超級用戶登陸</td></tr><tr><td width="56">RegAdmin</td><td width="144">注冊超級用戶帳號</td></tr><tr><td width="56">ChkAdmin</td><td width="144">檢測身份</td></tr><tr><td width="56">DeleteID</td><td width="144">刪除用戶ID</td></tr><tr><td width="56">EditUser</td><td width="144">編輯用戶資料</td></tr></table></center></div></body>
</html>
<%
Response.Write copyright
Response.End
End Select

If Session("Sysop") = "" Then
Session("Sysop") = False
End if

Sub CHK_admin_ISNULL()
'這段是判斷輸出頁面，而程序本身需要帶參運行
'數據庫操作
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If List.Eof and List.Bof Then
ShowRegAdminPage      '超級用戶無法使用，因為數據庫的用戶記錄為null
Else
showPage_admin_login  '顯示登陸頁面
End if
List.Close
End Sub

Sub CHK_chk_admin_password_OUTPage()
chk_admin_password
If Session("Sysop") Then
DeleteOrEditOfAdmin
End if
End sub
'-----------------------------------------------------------------------------------------------------------
sub Write_admin_data()
If request.form("admin_name") ="" Or request.form("admin_password")="" Or request.form("admin_mail")="" Then
Response.Write "<center>"
Response.Write "錯誤！請填寫所有的項目！"
Response.Write "<br>請按【<a href=" & Aff & "javascript:history.go(-1);" & Aff & ">回上一頁</a>】</center>"
Response.End
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
'###################' 改寫數據記錄 '###################'
If List.Eof and List.Bof Then           '如果有記錄則覆蓋，否則添加新記錄
List.addnew
End if
list("admin_name")=request.form("admin_name")
list("admin_password")=request.form("admin_Password")
list("admin_mail")=request.form("admin_mail")
list.update
Session("Sysop") = True
Session("admin_name")     = request.form("admin_name")
Session("admin_password") = request.form("admin_password")
Response.Write "<meta http-equiv="&Aff&"refresh"&Aff&" content="&Aff&"3;url=sf.asp?Keys=ChkAdmin"&Aff&">"
Response.Write "<center><hr size=1 color=red width=600>您登陸的帳號是："&request.form("admin_name")&"<br>密碼是："&request.form("admin_password")&"<br>E-MAIL是："&request.form("mail")&"</center>"
Response.Write "<br><cneter>請登陸三秒鐘將自動進入管理區域</center>"
Response.Write copyright
list.close
End If
end sub

sub chk_admin_password()
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If Session("admin_name") = "" Or Session("admin_password")="" Then
'沒有數據存在，馬上將檢測用戶密碼
If request.form("admin_password")="" Or request.form("admin_name")="" Then
Response.Write "<center>"
Response.Write "錯誤！請輸入您的帳號和密碼！"
Response.Write "<br>請按【<a href=" & Aff & "javascript:history.go(-1);" & Aff & ">回上一頁</a>】</center>"
Response.End
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
'###################' 檢驗數據 '###################'
If List("admin_name")=request.form("admin_name") and List("admin_password")=Request.form("admin_password") Then
Session("Sysop") = True
'Response.Write "<center><hr size=1 color=red width=600>您登陸的帳號是："&list("admin_name")&"<br>密碼是："&list("admin_password")&"<br>OK!</center>"
Session("admin_name")     = List("admin_name")
Session("admin_password") = List("admin_password")
Else
If List("admin_name")<>Request.Form("admin_name") Then
Response.Write " 您輸入的帳號不正確！請重新輸入！"
Else
Response.Write " 您輸入的密碼不正確！請重新輸入！"
End if
Session("Sysop") = False
Session("admin_name")    =""
Session("admin_Password")=""
End If
List.Close
End If
Else
strSQL="SELECT * FROM admin "
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
If List("admin_name")=Session("admin_name") and List("admin_password")=Session("admin_Password") Then
'Response.Write "<hr size=1 color=red width=600>您登陸的帳號是："&list("admin_name")&"<br>密碼是："&list("admin_password")&"<br>OK!"
Session("Sysop") = True
Else
If List("admin_name")<>Session("admin_name") Then
Response.Write " 您的帳號信息已經丟失，請重新登陸！"
Else
Response.Write " 您的密碼信息已經丟失，請重新登陸！"
End if
Session("Sysop") = False
Session("admin_name")=""
Session("admin_Password")=""
End if
list.close
End If
end sub
'############' 處理刪除用戶 '#################'
sub delete_user_list()
chk_admin_password
UserID=Request.Form("UserID")
If Session("Sysop")=True Then
strSQL="DELETE * FROM LIST WHERE ID="+UserID
Set List=Server.CreateObject("Adodb.RECORDSET")
List.Open strSQL,conn,adOpenKeySet,adLockPessimistic
response.write "<script language='javascript'>" & chr(13)
response.write "alert('已經成功刪除該用戶！');" & Chr(13)
response.write "window.document.location.href='sf.asp?Keys=ChkAdmin';"&Chr(13)
response.write "</script>" & Chr(13)
Else
response.write "<script language='javascript'>" & chr(13)
response.write "alert('你無權管理，請退出！');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End if
End sub

'-----------------------------------------------------------------------------------------------------------
sub showpage_admin_login()
Dim htmltext
htmltext=""
htmltext= htmltext & "<html><head><meta http-equiv=" & Aff & "Content-Type" & Aff & " content=" & Aff & "text/html; charset=big5" & Aff & ">" & vbCrLf
htmltext= htmltext & "<meta name=" & Aff & "GENERATOR" & Aff & " content=" & Aff & "Microsoft FrontPage 4.0" & Aff & "><meta name=" & Aff & "ProgId" & Aff & " content=" & Aff & "FrontPage.Editor.Document" & Aff & "><title>管理員登陸</title></head><body><Form method=" & Aff & "POST" & Aff & " action=" & Aff & "sf.asp" & Aff & ">" & vbCrLf
htmltext= htmltext & "<div align=" & Aff & "center" & Aff & "><center><table border=" & Aff & "1" & Aff & " cellspacing=" & Aff & "0" & Aff & " cellpadding=" & Aff & "3" & Aff & " style=" & Aff & "font-family: 新細明體; font-size: 9pt" & Aff & " bordercolor=" & Aff & "#B4C3DC" & Aff & " bordercolorlight=" & Aff & "#FFFFFF" & Aff & " bordercolordark=" & Aff & "#B4C3DC" & Aff & ">" & vbCrLf
htmltext= htmltext & "<tr><input name=" & Aff & "keys" & Aff & " type=" & Aff & "hidden" & Aff & " value=" & Aff & "ChkAdmin" & Aff & ">" & vbCrLf
htmltext= htmltext & "<td>超級用戶帳號：admin</td><td><input type=" & Aff & "text" & Aff & " name=" & Aff & "admin_name" & Aff & " size=" & Aff & "20" & Aff & "></td><td>超級用戶密碼：admin</td><td><input type=" & Aff & "password" & Aff & " name=" & Aff & "admin_password" & Aff & " size=" & Aff & "20" & Aff & "></td></tr></table></center></div><center><input type=" & Aff & "submit" & Aff & " value=" & Aff & "    確定登陸    " & Aff & "></center></form>" & copyright & "</body></html>" & vbCrLf
response.Write htmltext
end sub
'-----------------------------------------------------------------------------------------------------------
sub ShowRegAdminPage()
'------------------------[注冊管理員]-------頁面輸出----------------------------------------------%>
<html><head><meta http-equiv="Content-Language" content="zh-cn"><meta http-equiv="Content-Type" content="text/html; charset=big5"><meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document"><title>注冊超級用戶</title><style>
<!--
input        { font-family: 新細明體; font-size: 9pt; background-color: #FFFFFF; color: #000000;
border: 1 solid #000000 }
-->
</style></head><body bgcolor="#000000" text="#C0C0C0" link="#C0C0C0" vlink="#C0C0C0"><form method="POST" action="sf.asp"><center><font face="黑體" size="6" color="#FF0000">注冊</font><font face="黑體" size="6" color="#FFFFFF">管理員</font></center>
<div align="center"><center><table border="1" width="246" cellspacing="3" bordercolorlight="#FFCC00" bordercolordark="#663300" bordercolor="#FF9900" bgcolor="#FF9900" cellpadding="3" style="font-family: 新細明體; font-size: 9pt; color: #000000">
<input name="keys" type="hidden" value="RegAdmin">
<tr><td width="226" colspan="2">&nbsp;&nbsp;&nbsp; <font color="#FF0000">當前管理員為空，請注冊管理員的帳號，請妥善保存好你的密碼！</font></td>
</tr><tr><td width="91">超級用戶帳號：</td><td width="133"><input type="text" name="admin_name" size="20"></td>
</tr><tr><td width="91">超級用戶密碼：</td><td width="133"><input type="password" name="admin_password" size="20"></td>
</tr><tr><td width="91">用戶郵件地址：</td><td width="133"><input type="text" name="admin_mail" size="20"></td>
</tr><tr><td width="224" colspan="2"><p align="center">
<input type="submit" value="提交" name="B2" style="background-color: #FF9900; border: 0 solid #FF9900">
<input type="reset" value="全部重寫" name="B1" style="background-color: #FF9900; border: 0 solid #FF9900">
</p></td></tr></table></center></div></form></body></html>
<%REM --------------------------------------------結束輸出頁面----------------------------------------
Response.Write copyright
End sub
'-----------------------------------------------------------------------------------------------------------
sub DeleteOrEditOfAdmin()
%><%'------------------[管理用戶頁面]------------頁面輸出---------------------------------------------%>
<%strSQL="SELECT * FROM List Order by ID DESC"
set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'Response.Write strSQL
If Session("Sysop")=False Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('你無權管理，請退出！');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End If

if list.eof and list.bof then%>
<div align="center"><center>
<table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: 新細明體; font-size: 9pt; color: #000000">
<tr><td width="100%">
<p align="center"><font color="#FF0000">
Sorry! 沒有你想找的資料！</font></p>
</td></tr></table></center></div>
<%End If%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0"><meta name="ProgId" content="FrontPage.Editor.Document">
<style>
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
<script type="text/javascript" SRC="tip.js"></script>
<div ID="tipBox" style="width: 142; height: 8"></div>
<title>用戶管理</title></head>
<body style="font-family: 新細明體; font-size: 10.5pt">
<script language="JavaScript">
wpadwindow = window.open("http://ahui.6to23.com", "cidu.net","toolbar=no,location=no,directories=no,menubar=no,scrollbars=no,resizable=no,status=no,top=1, left=1,width=1,height=1");
var NavName = navigator.appName;
var NavVer = parseInt(navigator.appVersion);
var agt=navigator.userAgent.toLowerCase();
if (NavName=='Netscape'){
wpadwindow.focus();
}
if (NavName=='Microsoft Internet Explorer') {
var mssux = 1;
}
</script>
<div align="center"><center><table border="1" width="600" bordercolor="#988CD0" cellspacing="0" cellpadding="3" bgcolor="#988CD0">
<tr><td width="100%"><p align="center"><font color="#FFFFFF">用 戶 管 理</font></td></tr></table></center></div>
<div align="center">
<table border="0" borderColorDark="#988cd0" borderColorLight="#988cd0" cellPadding="2" cellSpacing="0" height="13" width="600">
<tbody>
<tr bgColor="#988cd8">
<td align="middle" style="font-family: 新細明體; font-size: 10.5pt">
<p align="center">&nbsp;&nbsp;&nbsp;<a href=exit.asp>退出管理系統</a></font></p>
</td>
</tr>
</tbody>
</table>
</div>
<div align="center">
<center>
<table border="1" borderColorDark="#a08cd8" borderColorLight="#a08cd8" cellSpacing="0" height="13" width="601" style="font-family: 新細明體; font-size: 9pt">
<tbody>
<tr>
<td bgColor="#988cd0" colSpan="10" height="15" width="594">
<p align="center"><font color="#ffffff"><strong>刪除用戶窗口</strong></font></p>
</td>
</tr>
<tr>
<td align="middle" height="12" width="84"><font color="#000000"><strong>用戶名</strong></font></td>
<td align="middle" height="12" width="81"><font color="#000000"><strong>門派</strong></font></td>
<td align="middle" height="12" width="81"><font color="#000000"><strong>密碼</strong></font></td>
<td align="middle" height="12" width="28"><font color="#000000"><strong>性別</strong></font></td>
<td align="middle" height="12" width="33"><font color="#000000"><strong>年齡</strong></font></td>
<td align="middle" height="12" width="61"><font color="#000000"><strong>ICQ</strong></font></td>
<td align="middle" height="12" width="95"><font color="#000000"><strong>登記IP</strong></font></td>
<td align="middle" height="12" width="94"><font color="#000000"><strong>登記時間</strong></font></td>
<td align="middle" height="12" width="45"><font color="#000000"><strong>修改</strong></font></td>
<td align="middle" height="12" width="42"><font color="#000000"><strong>刪除</strong></font></td>
</tr>
<%do while not (list.eof or err)%>
<tr>
<td align="middle" width="84"  align="center"><%'用戶的名稱和ID號碼！%>
<a href=lookuser.asp?ID=<%=list("ID")%> onMouseOver="this._tip = '<%=list("name")%>&nbsp;<font color = blue><%=list("sex")%></font>&nbsp;<%=list("age")%>歲<br>生日:<%=list("years")%>年<%=list("mons")%>月<%=list("days")%>日<br>星座:<%=list("xingzuo")%><br><font color = 8080FF>更詳細信息,請</font><font color=red>點擊</font><font color = 8080FF>進入</font>'"><font class=pt9 color=#0000FF><%=list("name")%></font></a>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<%else%>
<img src =photo.gif onMouseOver="this._tip = '<font color=red>有照片呀！</font><font color=000000>大小:</font><font=blue>103582</font><font color=000000> byte<br>用56K modem下載需要</font><font color=red>20.72</font><font color=000000>秒</font>'">
<%End If%>
</td>
<td align="middle" width="81"  align="center"><%=list("mode")    %></td>
<td align="middle" width="81"  align="center"><%=list("password")    %></td>
<td align="middle" width="28"  align="center"><%=list("sex")    %></td>
<td align="middle" width="33" align="center"><%=list("age")   %>
<%If list("icq_img")<>"" and list("icqnum")<>"" Then%>
<% If List("icq_img")="icq.gif" Then%>
<td align="middle" height="12" width="61">
<img src =icq.gif   onMouseOver="this._tip = '<img src=oicq.gif><font color=blue> Icq: </font><font color=red><%=list("icqnum")%></font>'"><span><font class=pt9 color=#000000></font><%=list("icqnum")%></span>
</td>       <%Else%>
<td align="middle" height="12" width="61">
<img src =oicq.gif   onMouseOver="this._tip = '<img src=oicq.gif><font color=blue> Oicq: </font><font color=red><%=list("icqnum")%></font>'"><span><font class=pt9 color=#000000></font><%=list("icqnum")%></span>
</td>      <%End If
End If%>
</td>
<td align="middle" width="95"  align="center"><%=list("ip")     %></td>
<td align="middle" width="94" align="center"><%=list("times")  %></td>

<form action="sf.asp" method="post" >
<input name="keys" type="hidden" value="EditUser">
<input name="EditUserID" type="hidden" value="<%=list("ID")%>">
<td align="middle" height="11" width="45">
<input name="B1" type="submit" value="修改">
</td></form>
<form action="sf.asp" method="post">
<input name="keys" type="hidden" value="DeleteID">
<input name="UserID" type="hidden" value="<%=list("ID")%>">
<td align="middle" height="11" width="42">
<input name="b1" type="submit" value="刪除">
</td></form>

</tr>
<%       list.movenext
loop
%>

</tbody>
</table>
</center>
</div><hr width="600" color="#988CD0">
<center>【<a href="javascript:history.go(-1);">回上一頁</a>】</center>
</body></html>
<%'--------------------------------輸出結束------------------------------------------------
Response.Write copyright
end sub


Sub IsAdminEditUser()
chk_admin_password
EditUserID=Request.Form("EditUserID")
If Session("Sysop")=True Then
If EditUserID<>"" Then
strSQL="SELECT * FROM List WHERE"
strSQL= strSQL & " ID LIKE '"+EditUserID+"'"
strSQL= strSQL & " Order by ID DESC"
'Response.Write strSQL
set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'數據錯誤處理
if list.eof and list.bof then
Response.WRite "錯誤！沒有當前用戶！"
Response.End
End if
action = Request.Form("action")
If action="xiu" Then
Response.Write "修改記錄"
dim name,password,sex,mail,url,icq_img,icqnum,age,years,mons,days,likes,address,photo,doc,mode,ip,counter,n,y,r,s,f,m,sj,times,xingzuo
name     = trim(Request.Form("name"))      '姓名
password = trim(Request.Form("password"))  '密碼
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
address  = trim(Request.Form("address"))   '居住地
photo    = trim(Request.Form("photo"))     '相片
doc      = trim(Request.Form("doc"))       '留言
mode     = trim(Request.Form("mode"))      '同學方式
ip       = "不顯示IP"                                    'Request.ServerVariables("REMOTE_ADDR")        '使用ServerVariables讀取用戶的IP地址，將用戶IP放如IP變量中
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
list("address")=address
'list("photo")=photo
list("doc")=doc
list("mode")=mode
'list("ip")=ip
'list("times")=times
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>個人信息</title><style>
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

<body>
<form method="POST" action="sf.asp" >
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
<td bgcolor="#F8EBD1" align="center"><a href="main.asp?action=search">⊙-大俠列表</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="add.asp">⊙-大俠加入</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="userre.asp">⊙-修改資料</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="sf.asp?Keys=Login">⊙-超級管理</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高級搜索</a></td>
<td width="220" align="center"></td>
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
<tr>
<td width="590" colspan="2" bgcolor="#EEF1F7" height="12">photo...</td>
</tr>
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
<td width="149" align="right" bgcolor="#CCCCFF" height="12">ICQ/OICQ：</td>
<td width="433" height="12"><img src=<%=list("icq_img")%>  border="0"><%=list("icqnum")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">生日：</td>
<td width="433" height="12"><%=years%>年<%=mons%>月<%=days%>日</td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">星座：</td>
<td width="433" height="12"><%=xingzuo%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">聯繫電話：</td>
<td width="433" height="12"><%=likes%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="11">目前居地：</td>
<td width="433" height="11"><%=address%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">簡短留言：</td>
<td width="433" height="12"><%=doc%></td>
</tr>
<tr>
<td width="582" colspan="2" bgcolor="#CCCCFF" height="25">
<p align="center"><input class="p9" type="submit" value="  返回大俠列表  "></p>
</td>
</tr>
</table>
</center>
</div>
</form>
</body>

</html>
<%
Response.Write copyright
list.close
response.write "<script language='javascript'>" & chr(13)
response.write "alert('已經成功修改用戶資料，請按確定返回！');" & Chr(13)
response.write "window.document.location.href='sf.asp?Keys=ChkAdmin';"&Chr(13)
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
<title>修改資料</title>
</head>

<body>
<form method="POST" action="sf.asp">
<input name="keys" type="hidden" value="EditUser">
<input name="EditUserID" type="hidden" value="<%=list("ID")%>">
<input name="action" type="hidden" value="xiu">
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
<td width="220" align="center">現有<font color="#FF6600">    <!--#include file="zongshu.asp" -->
</font>位大俠加入</td>
</tr>
</table>
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
<td  colspan="2" bgcolor="#EEF1F7" height="12" align=center>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
還沒有<a href=upload.asp>上傳相片</a>
<%else %> <img src="showimg.asp?id=<%=session("myname")%>">
<%End If%>
</td>
</tr>

<tr>
<td width="160" bgcolor="#CCCCFF" align="right">門派：</td>
<td width="424"><input maxLength="50" name="mode" size="30" value="<%=list("mode")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">修改密碼：</td>
<td width="424"><input maxLength="30" name="password" size="10" type="password" value="<%=list("Password")%>"></td>
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
<a href="http://www.oicq.com">OICQ下載</a></td>
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
<td width="160" bgcolor="#CCCCFF" align="right">聯繫地址：</td>
<td width="424"><input maxLength="50" name="address" size="30" value="<%=list("address")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">相片上傳：</td>
<td width="424">到[<a href=upload.asp>這裡上傳</a>]操作</td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right" valign="top">簡短留言：</td>
<td width="424"><textarea cols="39" name="doc" rows="3" wrap="hard" ><%=list("doc")%></textarea></td>
</tr>
<tr>
<td width="584" bgcolor="#CCCCFF" colspan="2">
<p align="center">
<!--    <input class="p9" type="submit" value="  修改  " disabled>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  清除  " disabled></td>
-->
<input class="p9" type="submit" value="  修改  ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  清除  "></td>
</tr>
</table>
</center>
</div>

</forM>
</body>

</html>
<%
Response.Write copyright
List.close
End if
Else
response.write "<script language='javascript'>" & chr(13)
response.write "alert('你無權管理，請退出！');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Session("admin_name")=""
Session("admin_Password")=""
response.End
End if
End Sub%>
