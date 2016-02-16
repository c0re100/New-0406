
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>迎花宮</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#660000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<br>
<table border=1 bgcolor="#FFFFFF" align=center width=80% cellpadding="0" cellspacing="1" bordercolor="#FFCC00">
  <tr bgcolor="#669900"> 
    <td height="16" colspan="5"> 
      <p align=center class="p9"><font color="#FF3333"><b><font color="#FFFFFF">靈劍江湖迎花宮</font></b></font></p>
    
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="14" bgcolor="#FFFFCC" bordercolor="#FF6600" width="30%"> 
      <div align="center"><font size="2" color="#FF0000">帥 男 酷 號</font></div>
    </td>
    <td height="14" bgcolor="#FFFFCC" bordercolor="#FF6600" width="10%"> 
      <div align="center"><font size="2" color="#FF0000">帥 氣</font></div>
    </td>
    <td height="14" width="15%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">服 務 價 格</font></div>
    </td>
    <td height="14" width="30%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">增 加 的 內 力</font></div>
    </td>
    <td height="14" width="15%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">操 做</font></div>
    </td>
  </tr>
  <!--#include file="jiu.asp"-->
  <%
sql="SELECT id ,姓名,美貌度 FROM 帥男"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor=#DEAD63> 
    <td bgcolor="#FFFFFF" width="30%"> 
      <center>
        <font size="2"><%=rs("姓名")%><span class="p9"><font color="#6666FF"></font></span> 
        </font> 
      </center>
      <font size="2"></font> 
    <td bgcolor="#FFFFFF" width="10%"> 
      <div align="center"><font size="2"><%=rs("美貌度")*0.1%></font> </div>
    
    <td bgcolor="#FFFFFF" width="15%"> 
      <div align="center"><font size="2"><%=rs("美貌度")*0.5%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="30%"> 
      <div align="center"><font size="2"><%=rs("美貌度")*0.8%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="15%"> 
      <center>
        <font size="2"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">進入狀態</span></a></font> 
      </center>
    </td>
  </tr>
  <%
rs.movenext
loop
set rs=nothing
connt.close%>
</table>
</BODY></HTML>     
