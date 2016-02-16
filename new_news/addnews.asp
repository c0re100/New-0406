<%
If Request("new") <> Empty Then
d = Request.Form("d")
subject = Request.Form("subject")
title = Request.Form("title")
name = Request.Form("name")
news = Request.Form("news")
news = Replace(news,vbCrLf , "<br>")

Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "select * from news ORDER BY id DESC", conn,1,3

rs.AddNew
rs("d") = d
rs("subject") = subject
rs("title") = title
rs("name") = name
rs("news") = news
rs.update
rs.close
SET CONN = NOTHING
Session("Login") = True
%>
<script language="javascript">
	alert("公告已經新增完畢.返回管理頁!")
	location.href="admin-home.asp"
</script>

<% ELSE %>
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
<title>新增公告</title>
<!--#INCLUDE FILE="news.asp" --></head>
  <center><a href="admin-home.asp"><font size="2">回管理中心</font></a>
  <form action="addnews.asp" method="post" name="add">
  <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%"><font size="2">新增公告</font></td>
            </center>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <center>
  <tr>
    <td width="300">
      <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="300">
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">日期：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="d" size="20" value="<%=date%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">公告主旨：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="subject" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">公告標題：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="title" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">發表人</font><font size="2">：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="name" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" width="298" bgcolor="#FFFFFF"><font size="2">詳細內容：</font></td>    
          </tr>
          <tr>
        <td colspan="2" height="98" width="298" bgcolor="#FFFFFF">     
          <font size="2">     
          <textarea name="news" rows="4"  cols="47"></textarea>    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" align=center bgcolor="#E1E1C4" width="298">     
          <font size="2">     
          <input type="submit" name="new" value="確定送出">    
          <input type="reset" value="重新設定" id="reset1" name="reset1"    
          </font>     
        </td>     
          </tr> 
                  </table> 
          </td> 
  </tr> 
  </table> 
  </center> 
</form>   
</html>
<%
end if
%>