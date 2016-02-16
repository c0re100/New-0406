<% 
If Request("UP") <> Empty Then
id = Request("id")
d=Request.Form("d")
subject=Request.Form("subject")
title=Request.Form("title")
news=Request.Form("news")

Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sqlstr="select * from news where id=" & id
rs.open sqlstr,conn,3,2
rs("d")=d
rs("subject")=subject
rs("title")=title
rs("news")=news
rs.update
rs.close
set conn=nothing
Session("Login") = True
%>
<script language="javascript">
	alert("修改公告完畢.返回管理頁!")
	location.href="admin-home.asp"
</script>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>修改公告</title>
<%
id=request("id")
set adocon=Server.CreateObject("ADODB.Connection")
adocon.Open "Driver={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("new.mdb")
SQL="SELECT * FROM news WHERE id = "& id &""
Set rs = adocon.Execute(SQL)
do while not rs.eof
%>
<!--#INCLUDE FILE="news.asp" -->
</head>
  <center>
  <form action="up.asp?id=<% =id %>" name="form" method="post">
  <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%"><font size="2">修改公告</font></td>
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
          <input type="text" name="d" size="20" value="<%=rs("d")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">公告主旨：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="subject" size="20" value="<%=rs("subject")%>">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">公告標題：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="title" size="20" value="<%=rs("title")%>">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" width="298" bgcolor="#FFFFFF"><font size="2">詳細內容：</font></td>    
          </tr>
          <tr>
        <td colspan="2" height="98" width="298" bgcolor="#FFFFFF">     
          <font size="2">     
          <textarea name="news" rows="4"  cols="47">
<%
	content=replace(rs("news"),"<br>",chr(13))
	content=replace(content,"&nbsp;"," ")
	response.write content
%>
</textarea>    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" align=center bgcolor="#E1E1C4" width="298">     
          <font size="2">     
          <input type="submit" name="UP" value="確定修改">    
         <input type="reset" value="重新設定" id="reset1" name="reset1">
                   </font>     
          <%rs.movenext
loop
%>
        </td>     
          </tr> 
                  </table> 
          </td> 
  </tr> 
  </table> 
  </center> 
</form>   
</html>
subject