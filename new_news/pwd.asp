<%
If Request.QueryString("logout")="true" Then
Session.Abandon
Response.Redirect "new.asp"
End If
%>


<%
If Request("pwd") <> Empty Then
name=request("name")
pwd=request("pwd")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("new.mdb")
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * From admin where name='"& name &"' and pwd= '"& pwd &"'"
set rs=conn.execute(sql)
%>
<%
if rs.eof then
%>
<script language="javascript">
	alert("警告：帳號密碼錯誤")
	location.href="javascript:history.go(-1)"
</script>
<%
else
 	session("name")=rs("name")
	session("pwd")=rs("pwd")
	Session("lv")= rs("lv")
	session("AdminLogin")=True
	response.redirect "admin-home.asp"
end if
%>
<% ELSE %>
<%
if session("name")<>"" and session("pwd")<>"" then
response.redirect "admin-home.asp"
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>管理模式</title>
</head>
<!--#INCLUDE FILE="news.asp" -->

<body>

<div align="center">
  <center>
  <table border="1" cellpadding="3" cellspacing="0" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#B7D3FF" bordercolordark="#E1E1C4" align="center"><font size="3">登入管理</font></td>
    </tr>
    <tr><form method="POST" action="pwd.asp">
      <td>
        
          <p align="center">帳號：<input type="text" name="name" size="20" style="font-size: 9pt"><br>
          密碼：<input type="password" name="pwd" size="20" style="font-size: 9pt"><br>
          <input type="submit" value="登錄" name="B1" style="width: 40; height: 19; font-size: 9pt; background-color: #dddddd; background-repeat: repeat; background-attachment: scroll; border: 1px groove #000000; background-position: 0% 50%">
          <input type="reset" value="取消" name="B2" style="width: 40; height: 19; font-size: 9pt; background-color: #dddddd; background-repeat: repeat; background-attachment: scroll; border: 1px groove #000000; background-position: 0% 50%"></p> 
         
      </td> 
    </tr></form> 
  </table> 
  </center> 
</div> 
 
</body> 
 
</html> 
<% 
End If 
%>