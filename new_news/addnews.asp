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
	alert("���i�w�g�s�W����.��^�޲z��!")
	location.href="admin-home.asp"
</script>

<% ELSE %>
<%
if session("AdminLogin") = empty then
%>
<script language="javascript">
	alert("ĵ�i�G�Ф��|�իD�k�n�J")
	location.href="pwd.asp"
</script>
<%
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�s�W���i</title>
<!--#INCLUDE FILE="news.asp" --></head>
  <center><a href="admin-home.asp"><font size="2">�^�޲z����</font></a>
  <form action="addnews.asp" method="post" name="add">
  <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%"><font size="2">�s�W���i</font></td>
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
        <td width="124" bgcolor="#FFFFFF"><font size="2">����G</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="d" size="20" value="<%=date%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">���i�D���G</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="subject" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">���i���D�G</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="title" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">�o��H</font><font size="2">�G</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="name" size="20">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" width="298" bgcolor="#FFFFFF"><font size="2">�ԲӤ��e�G</font></td>    
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
          <input type="submit" name="new" value="�T�w�e�X">    
          <input type="reset" value="���s�]�w" id="reset1" name="reset1"    
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