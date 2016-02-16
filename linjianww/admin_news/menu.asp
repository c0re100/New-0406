<!--#include file="../inc/global.news.asp" -->
<TABLE border="0" width="100%" cellpadding="0" cellspacing="1">
  <TR> 
    <TD bgcolor="#4880a8" height="20" align="center"><FONT color="#FFFFFF">管 理 
      選 項</FONT></TD>
  </TR>
  <%
j=UBound(arr_news_type)
For i=1 to j
	if trim(arr_news_type(i,1))<>"" then 
%>
  <TR> 
    <TD height="20" bgcolor="#e7f0f8" align="center" style="cursor : hand;" onmouseover="this.bgColor='#b8d4e8';" onmouseout="this.bgColor='#e7f0f8';"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="170">&nbsp;</td>
          <td width="831"> 
            <%if arr_news_type(i,2)>"1" then response.write "　"%>
            <img src="../image/arrow.gif" width="5" height="5"> 
            <%if arr_news_type(i,2)="1" then %>
            <%=arr_news_type(i,1)%> 
            <%else%>
            <a href="news_list.asp?type=<%=i%>"><%=arr_news_type(i,1)%></a> 
            <%end if%>
          </td>
        </tr>
      </table>
      <%
	end if 
Next
%>
    </TD>
  </TR>
</TABLE>
