<%
If Request("nuse") <> Empty Then
name = Request.Form("name")
pwd = Request.Form("pwd")
lv = Request.Form("lv")

Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "select * from admin ORDER BY id DESC", conn,1,3
rs.AddNew
rs("name") = name
rs("pwd") = pwd
rs("lv") = lv
rs.update
rs.close
SET CONN = NOTHING
Session("Login") = True
%>
<script language="javascript">
	alert("管理員已經新增完畢.返回管理頁!")
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
<title>新增管理員</title>
<!--#INCLUDE FILE="news.asp" -->
</head>
  <center>
  <form action="nuse.asp" method="post" name="add">
    <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%">
                <p align="center"><font size="2">新增管理員</font></p>
              </td>
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
        <td width="124" bgcolor="#FFFFFF"><font size="2">登入名稱：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="name" size="20" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">登入密碼：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="pwd" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">管理等級：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <select size="1" name="lv">
          <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          </select><font size="2">     
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" align=center bgcolor="#E1E1C4" width="298">     
          <font size="2">     
          <input type="submit" name="nuse" value="確定送出">    
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