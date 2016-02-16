<% 
If Request("id") <> Empty Then 
wed = Request("wed")
page=Request.Form("page")
cor=Request.Form("Cor")
cor1=Request.Form("Cor1")
cor2=Request.Form("Cor2")
cor3=Request.Form("Cor3")
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sqlstr="select * from wed where wed"
rs.open sqlstr,conn,3,2
rs("cor")=cor
rs("cor1")=cor1
rs("cor2")=cor2
rs("cor3")=cor3
rs("wed")=wed
rs("page")=page
rs.update
rs.close
set conn=nothing
Session("Login") = True
%>
<script language="javascript">
	alert("基本環境已經設定完畢.返回管理頁!")
	location.href="admin-home.asp"
</script><% ELSE %>
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
<title>基本環境設定</title>
<%
set adocon=Server.CreateObject("ADODB.Connection")
adocon.Open "Driver={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("new.mdb")
SQL="SELECT * FROM wed WHERE wed "

Set rs = adocon.Execute(SQL)
do while not rs.eof
%>
<!--#INCLUDE FILE="news.asp" -->
</head>
  <center><a href="admin-home.asp"><font size="2">回管理中心</font></a>
  <form action="wed.asp" name="form" method="post">
  <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300">
        <div align="center">
          <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
              <td width="100%"><font size="2">基本環境設定</font></td>
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
        <td width="124" bgcolor="#FFFFFF"><font size="2">標題欄背景顏色</font><font size="2">：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="cor" size="20" value="<%=rs("cor")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">公告欄顏色</font><font size="2">：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="cor1" size="20" value="<%=rs("cor1")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">分頁欄</font><font size="2">背景顏色：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="cor2" size="20" value="<%=rs("cor2")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">版權欄背景顏色：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="cor3" size="20" value="<%=rs("cor3")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">網站名稱：</font></td>   
        <td width="172" bgcolor="#FFFFFF">   
          <font size="2">   
          <input type="text" name="wed" size="20" value="<%=rs("wed")%>" style="color:blue">   
          </font>   
        </td>   
          </tr>
          <tr>
        <td width="124" bgcolor="#FFFFFF"><font size="2">每頁筆數：</font></td>    
        <td width="172" bgcolor="#FFFFFF">     
          <font size="2">     
          <input type="text" name="page" size="20" value="<%=rs("page")%>">    
          </font>    
        </td>    
          </tr>
          <tr>
        <td colspan="2" align=center bgcolor="#E1E1C4" width="298">     
          <font size="2">     
          <input type="submit" name="id" value="確定修改">    
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
<% 
End If 
%> 























