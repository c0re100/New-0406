<% 
If Request("use") <> Empty Then
id = Request("id")
name=Request.Form("name")
pwd = Request("pwd")
lv = Request("lv")

Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sqlstr="select * from admin where id=" & id
rs.open sqlstr,conn,3,2
rs("name")=name
rs("pwd")=pwd
rs("lv")=lv
rs.update
rs.close
set conn=nothing
Session("Login") = True
%>
<script language="javascript">
	alert("�ޭ���Ƥw�g��s����.��^�޲z��!")
	location.href="admin-home.asp"
</script>

<% END IF %>

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
<title>�ק�޲z�����</title>
<%
id=request("id")
set adocon=Server.CreateObject("ADODB.Connection")
adocon.Open "Driver={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("new.mdb")
SQL="SELECT * FROM admin WHERE id = "& id &""
Set rs = adocon.Execute(SQL)
do while not rs.eof
%>
<!--#INCLUDE FILE="news.asp" -->
</head>
  <center>
  <form action="uses.asp?id=<% =id %>" name="form" method="post">
  <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%">
                <p align="center"><font size="2">�ק�޲z�����</font></p>
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
        <td width="124" bgcolor="#FFFFFF"><font size="2">�ק�ޭ��b���G</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="name" size="20" value="<%=rs("name")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">�ק�ޭ��K�X�G</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="password" name="pwd" size="20" value="<%=rs("pwd")%>">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">�ק�ޭ����šG</font></td>    
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
          <input type="submit" name="use" value="�T�w�ק�">    
          <input type="reset" value="���s�]�w" id="reset1" name="reset1">     
          </font>
        </td>     
          </tr> 
                  </table> 
          </td> 
  </tr> 
  </table> 
  </center> 
 </form>  
<%rs.movenext
loop %></html>