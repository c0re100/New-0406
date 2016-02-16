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
<title>管理員清單</title>
<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SortSql = "Select *From admin order by id Desc"
rs.Open SortSql, conn, 1,3
%>
<!--#INCLUDE FILE="news.asp" -->
</head>
<center>
<font size="2">
<a href="nuse.asp">新增管理員</a> <a href="admin-home.asp">回管理中心</a></font><BR> 
    <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
    <tr> 
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300"> 
        <div align="center"> 
          <table border="0" cellpadding="3" cellspacing="0" width="100%"> 
            <tr> 
              <td width="100%"> 
                <p align="center"><font size="2">管理員清單</font></p> 
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
        <td width="25%" bgcolor="#FFFFFF" align="center"> 
          <p align="center"><font size="2">管員</font><font size="2">帳號</font></p> 
        </td>    
        <td width="25%" bgcolor="#FFFFFF" align="center"><font size="2">管理等級</font></td>    
        <td width="25%" bgcolor="#FFFFFF" align="center">    
          <font size="2">    
          刪除管員    
          </font>    
        </td>    
        <td width="25%" bgcolor="#FFFFFF" align="center">    
          <font size="2">修改資料</font>    
        </td>    
          </tr> 
          <%FOR SH=1 to RS.PageSize                                            
id=rs("id")                                            
%>   
          <tr> 
        <td width="25%" bgcolor="#FFFFFF" align="center"><font size="2"><%=rs("name")%></font></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><%=rs("lv")%></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><a href="delns.asp?id=<%=rs("id")%>">刪除</a></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><a href="uses.asp?id=<%=rs("id")%>">修改</a></td>     
         </tr> 
         <%          
rs.MoveNext          
IF RS.EOF THEN EXIT FOR          
Next             
%>    
          <tr> 
        <td colspan="4" align=center bgcolor="#E1E1C4" width="298">      
          <font size="2"></font> 
        </td>      
          </tr>  
                  </table>  
          </td>  
  </tr>  
  </table>  
  </center>   
</html> 
 <%rs.close                   
conn.close                
%>    
 
 
 
 
 
 
 
 
 
 
 
 
