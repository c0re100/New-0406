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
<title>�޲z���M��</title>
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
<a href="nuse.asp">�s�W�޲z��</a> <a href="admin-home.asp">�^�޲z����</a></font><BR> 
    <table border="1" cellpadding="0" cellspacing="0" width="280" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
    <tr> 
      <td bgcolor="#E1E1C4" bordercolordark="#E1E1C4" width="300"> 
        <div align="center"> 
          <table border="0" cellpadding="3" cellspacing="0" width="100%"> 
            <tr> 
              <td width="100%"> 
                <p align="center"><font size="2">�޲z���M��</font></p> 
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
          <p align="center"><font size="2">�ޭ�</font><font size="2">�b��</font></p> 
        </td>    
        <td width="25%" bgcolor="#FFFFFF" align="center"><font size="2">�޲z����</font></td>    
        <td width="25%" bgcolor="#FFFFFF" align="center">    
          <font size="2">    
          �R���ޭ�    
          </font>    
        </td>    
        <td width="25%" bgcolor="#FFFFFF" align="center">    
          <font size="2">�ק���</font>    
        </td>    
          </tr> 
          <%FOR SH=1 to RS.PageSize                                            
id=rs("id")                                            
%>   
          <tr> 
        <td width="25%" bgcolor="#FFFFFF" align="center"><font size="2"><%=rs("name")%></font></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><%=rs("lv")%></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><a href="delns.asp?id=<%=rs("id")%>">�R��</a></td>     
        <td width="25%" bgcolor="#FFFFFF" align="center"><a href="uses.asp?id=<%=rs("id")%>">�ק�</a></td>     
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
 
 
 
 
 
 
 
 
 
 
 
 
