<!--#INCLUDE FILE="conn.asp" -->
<%
PAGE = Request("PAGE")
IF PAGE = "" THEN
  PAGE = 1
ELSE
  PAGE = PAGE
END IF
RS.PageSize = 10
If Not rs.eof Then
   RS.AbsolutePage = PAGE
End if
%>
                                                
                                 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>���򤽧i</title>
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
                                                                                



<center><table border="1" borderColor="#111111" cellPadding="2" cellSpacing="0" style="BORDER-COLLAPSE: collapse" width="330" height="38">  
<tr> 


<td bgColor="<%=rs1("cor")%>" width="60" height="7" align="center">  
<p align="center"><font color="#000000" size="2">�C��<font color="#FF0000" size="2"><%=rs1("page")%></font><font color="#000000" size="2">��</font>  
</p>   
</td>   
<td bgColor="<%=rs1("cor")%>" width="210" height="7" align="center">  
<p align="center"><font color="#000000" size="2">  
<%=rs1("wed")%>new~~0406���򤽧i</font>    
</td>    
<td bgColor="<%=rs1("cor")%>" width="60" height="7" align="center">    
<p align="center"><font size="2"><a href="pwd.asp" target="_blank">�n�J�޲z</a></font>  
</td>  
</tr> 
      <IMG src="img/jh_pic.gif" width="327" height="22" border="0">
<tr><td width="322" height="19" colspan="3" bgcolor="<%=rs1("cor1")%>"> 
<p align="center"> 
<% if rs.EOF or rs.BOF then %>
<font size="2">�ثe�ȵL�o��s�D</font> 
<% else %>    
<%FOR SH=1 to RS.PageSize                                           
news=rs("news")                                           
%>           
<table border="0" borderColor="#000000" cellPadding="2" cellSpacing="0" style="BORDER-COLLAPSE: collapse" width="100%">      
<tr>      
<td width="4%" bgcolor="#FFFFFF"><img border="0" src="img/00.gif"></td>      
<td width="76%" bgcolor="#FFFFFF"><font size="2" color="#990033"><b><%=rs("subject")%>:</b></font><font size="2">  <a href="javascript:openWindows('seenews.asp?id=<%=rs("id")%>','','scrollbars=yes,top=100,left=300,width=575,height=275')"><%=rs("title")%></a></font><font color="#CC6600"><%if year(rs("d"))=year(date()) and month(rs("d"))=month(date()) and day(rs("d"))=day(date()) then%></font><img src="img/new.gif" align="absmiddle"><%end if%></td>            
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
<td bgColor="<%=rs1("cor2")%>" width="330" height="7" COLSPAN=3>              
<p align="center">         
         
<font size="2">         
         
<%                                                                
   If Page <> 1 Then                                                               
      Response.Write "<A HREF=new.asp?Page=1><font size=2>�Ĥ@��</font></A>�@"                                                               
      Response.Write "<A HREF=new.asp?Page=" & (Page-1) & "><font size=2>�W�@��</font></A>�@"                                                               
   End If                                                               
   If Page <>  rs.PageCount Then                                                               
      Response.Write "<A HREF=new.asp?Page=" & (Page+1) & "><font size=2>�U�@��</font></A>�@"                                                               
      Response.Write "<A HREF=new.asp?Page=" &  rs.PageCount & "><font size=2>�̫�@��</font></A>�@"                                                               
   End If                
%>  <font color=#7102AC size=2>����:</font><FONT COLOR="Red" size=2><%=Page%>/<%= rs.PageCount%></FONT>                      
</font>                      
</td>                      
</tr>                    
<tr> 
                 
<td bgColor="<%=rs1("cor3")%>" width="330" height="7" COLSPAN=3>                        
<p align="center"><font color="#000000" size="2">�{���ק�/�]�p:���</a>  <BR>
���v�޶i:new~~0406���򤽧i�t��       </font>                   
</td>                   
</tr>                  
</table>             
<% End if %><%rs.close                    
conn.close                 
%>