<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
%>
<HTML><HEAD><TITLE>�F�C����-�ɬ��|</TITLE><META content="text/html; charset=big5" http-equiv=Content-Type><LINK
href="../new.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD><BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">::<font size="2" color="#FD794D"><b>�ɬ��|</b></font>::<br><br>
</div>
<table border=1 bgcolor="#FFFFFF" align=center width=565 cellpadding="0" cellspacing="1" bordercolor="#ff7010">
  <tr bgcolor="#FFFFFF"> 
 
     
    
  <tr bgcolor=#FFFFFF bordercolor="#000000"> 
    <td height="15" bordercolor="#FF6600" colspan="7"> 
      <div align="center"><font color="#FF3333"><b>�F�C����ɬ��|</b></font></div>
      </td>
  </tr>
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="57"> 
      <div align="center"><font size="2" color="#FF0000">�W��</font></div>
    </td>
    <td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="62"> 
      <div align="center"><font size="2" color="#FF0000">����</font></div>
    </td>
    <td height="15" width="107" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�M���Ѫ�����</font></div>
    </td>
    <td height="15" width="63" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">���ФH</font></div>
    </td>
    <td height="15" width="107" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�M���ѼW�[����O</font></div>
    </td>
    <td height="15" width="87" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�ө�</font></div>
    </td>
    <td height="15" width="58" bgcolor="#FFFFCC" bordercolor="#FF6600"><font size="2" color="#FF0000">�@�Ϧ��</font></td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �W��"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor=#DEAD63> 
    <td bgcolor="#FFFFCC" width="57"> 
      <center>
        <font size="2"><%=rs("�m�W")%><span class="p9"><font color="#6666FF"></font></span></font> 
      </center>
      <font size="2"></font> 
    <td bgcolor="#FFFFCC" width="62"> 
      <div align="center"><font size="2"><%=rs("������")%></font> </div>
    <td bgcolor="#FFFFCC" width="107"> 
      <div align="center"><font size="2"><%=rs("������")*1.5%> </font></div>
    </td>
    <td bgcolor="#FFFFCC" width="63"> 
      <div align="center"><font size="2"><%=rs("����")%> </font></div>
    </td>
    <td bgcolor="#FFFFCC" width="107"> 
      <div align="center"><font size="2"><%=rs("������")/2%> </font></div>
    </td>
    <td bgcolor="#FFFFCC" width="87"> 
      <center>
        <font size="2"><a href=g.asp?id=<%=rs("id")%>><span class="calen-curr">�i�J��ө�</span></a></font> 
      </center>
    </td>
    <td bgcolor="#FFFFCC" width="58"> 
      <div align="center"><font size="2"><a href=zhengjiu.asp?id=<%=rs("id")%>><span class="calen-curr">�ϴ�</span></a></font></div>
    </td>
  </tr>
  <%
rs.movenext
loop
rs.close
	set rs=nothing
	conn.close
	set conn=nothing%>
</table>
</BODY></HTML> 