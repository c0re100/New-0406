<%
if session("AdminLogin") = empty then
%>
<script language="javascript">
	alert("警告：請勿嚐試非法登入")
	location.href="pwd.asp"
</script>
<%
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>管理中心</title>
<style>
<!--
a {text-decoration: none;color:#0000C0}
a:hover {text-decoration:underline;color:#ff9900}
// -->
</style>
</head>
<script language="JavaScript">
<!--
function openWindows(theURL,winName,features) {
  window.open(theURL,winName,features);
}
//-->
</script>
<div align="center">
  <center>
  <table border="1" cellpadding="3" width="300" bordercolor="#B7D3FF">
    <tr>
      <td bgcolor="#B7D3FF">
        <p align="center"><font color="#000000" size="2">管理中心</font></td>
    </tr>
    <%if Session("lv") =4 OR Session("lv") =3 OR Session("lv") =2 OR Session("lv") =1 then%>
    <tr>
      <td>
        <p align="center"><a href="addnews.asp"><font size="2" color="#FF0000">新增公告</font></a></td>
    </tr>
    <%end if%>
    <%if Session("lv") =4 OR Session("lv") =3 OR Session("lv") =2 OR Session("lv") =1 then%>

      <tr>
      <td>
        <p align="center"><a href="admin.asp"><font size="2" color="#FF0000">刪除.修改公告</font></a></td>
    </tr>
    <%end if%>
    <%if Session("lv") =4 OR Session("lv") =3 OR Session("lv") =2 OR Session("lv") =1 then%>
    <tr>
      <td>
        <p align="center"><a href="wed.asp"><font size="2" color="#FF0000">基本環境設定</font></a></td>
    </tr>
    <%end if%>
    <%if Session("lv") =4 OR Session("lv") =3 OR Session("lv") =2 OR Session("lv") =1 then%>
    <tr>
      <td>
        <p align="center"><a href="use.asp"><font size="2" color="#FF0000">新增.修改.刪除管理員</font></a></td>
    </tr>
    <%end if%>
    <tr>
      <td>
        <p align="center"><a href="pwd.asp?logout=true">
        <font size="2" color="#FF0000">登出</font></a></td>
    </tr>
  </table>
  </center>
</div>