
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>���c</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#660000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<br>
<table border=1 bgcolor="#FFFFFF" align=center width=80% cellpadding="0" cellspacing="1" bordercolor="#FFCC00">
  <tr bgcolor="#669900"> 
    <td height="16" colspan="5"> 
      <p align=center class="p9"><font color="#FF3333"><b><font color="#FFFFFF">�F�C������c</font></b></font></p>
    
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="14" bgcolor="#FFFFCC" bordercolor="#FF6600" width="30%"> 
      <div align="center"><font size="2" color="#FF0000">�� �k �� ��</font></div>
    </td>
    <td height="14" bgcolor="#FFFFCC" bordercolor="#FF6600" width="10%"> 
      <div align="center"><font size="2" color="#FF0000">�� ��</font></div>
    </td>
    <td height="14" width="15%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�A �� �� ��</font></div>
    </td>
    <td height="14" width="30%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�W �[ �� �� �O</font></div>
    </td>
    <td height="14" width="15%" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�� ��</font></div>
    </td>
  </tr>
  <!--#include file="jiu.asp"-->
  <%
sql="SELECT id ,�m�W,������ FROM �Өk"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor=#DEAD63> 
    <td bgcolor="#FFFFFF" width="30%"> 
      <center>
        <font size="2"><%=rs("�m�W")%><span class="p9"><font color="#6666FF"></font></span> 
        </font> 
      </center>
      <font size="2"></font> 
    <td bgcolor="#FFFFFF" width="10%"> 
      <div align="center"><font size="2"><%=rs("������")*0.1%></font> </div>
    
    <td bgcolor="#FFFFFF" width="15%"> 
      <div align="center"><font size="2"><%=rs("������")*0.5%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="30%"> 
      <div align="center"><font size="2"><%=rs("������")*0.8%> </font></div>
    </td>
    <td bgcolor="#FFFFFF" width="15%"> 
      <center>
        <font size="2"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">�i�J���A</span></a></font> 
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
