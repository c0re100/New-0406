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
<!--#INCLUDE FILE="conn.asp" -->
<%
PAGE = Request("PAGE")
IF PAGE = "" THEN
  PAGE = 1
ELSE
  PAGE = PAGE
END IF
RS.PageSize = 3
If Not rs.eof Then
   RS.AbsolutePage = PAGE
End if
%>
<html>
<head>
<title>刪除.修改公告</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
      
<p align="center"><a href="admin-home.asp"><font size="2">回管理中心</font></a><BR>
      
<center><table border="1" borderColor="#111111" cellPadding="2" cellSpacing="0" style="BORDER-COLLAPSE: collapse" width="330" height="38"> 
<tr>
<td bgColor="#99ccff" width="330" height="7">      
<p align="center"><font color="#000000" size="2">刪除.修改公告</font>   
</td>   
</tr>
<tr><td width="322" height="19">
<p align="center">
<% if rs.EOF or rs.BOF then %>
<font size="2">目前暫無資料可管理</font> 
<% else %>     
 
<%FOR SH=1 to RS.PageSize                                         
news=rs("news")                                         
%>         
<table border="0" borderColor="#111111" cellPadding="2" cellSpacing="0" style="BORDER-COLLAPSE: collapse" width="100%">    
<tr>    
<td width="4%" bgcolor="#FFFFFF"><img border="0" src="img/00.gif"></td>    
<td width="76%" bgcolor="#FFFFFF"><font size="2"><font size="2" color="#990033"><b><%=rs("subject")%>:</b></font><font size="2">  <a href="javascript:openWindows('seenews.asp?id=<%=rs("id")%>','','scrollbars=yes,top=100,left=300,width=350,height=200')"><%=rs("title")%></a></font><BR>  
<a href="deln.asp?id=<%=rs("id")%>">刪除</a>  <a href="up.asp?id=<%=rs("id")%>">修改</a>       
       
</font></td>          
<td width="20%" bgcolor="#FFFFFF"><font size="2"><%=rs("d")%></font></td>          
</tr>          
</table>          
<hr noshade size="0" color="#000000">       
<%          
rs.MoveNext          
IF RS.EOF THEN EXIT FOR          
Next             
%>          
         
<tr>        
<td bgColor="#99ccff" width="330" height="7">            
<p align="center">       
       
<%                                                              
   If Page <> 1 Then                                                             
      Response.Write "<A HREF=admin.asp?Page=1><font size=2>第一頁</font></A>　"                                                             
      Response.Write "<A HREF=admin.asp?Page=" & (Page-1) & "><font size=2>上一頁</font></A>　"                                                             
   End If                                                             
   If Page <>  rs.PageCount Then                                                             
      Response.Write "<A HREF=admin.asp?Page=" & (Page+1) & "><font size=2>下一頁</font></A>　"                                                             
      Response.Write "<A HREF=admin.asp?Page=" &  rs.PageCount & "><font size=2>最後一頁</font></A>　"                                                             
   End If              
%>  <font color=#7102AC size=2>頁次:</font><FONT COLOR="Red" size=2><%=Page%>/<%= rs.PageCount%></FONT>           
</td>           
</tr>         
</table>       
<%      
End if      
%>         
 <%rs.close           
conn.close        
%>        
</center>