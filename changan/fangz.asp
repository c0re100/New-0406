<% if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM �Ы� where  ����='�@��Ы�'"
Set Rs=conn.Execute(sql)
name=session("myname")
%>
<html>
<LINK 
href="../pic/css.css" rel=stylesheet>
<body bgcolor="#FFFFFF">
<table border=1 bgcolor="#FFFFFF" align=center width=54% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099">�w��</font><%=name%><font color="#990099">���{�в����ߡA�ֳ��Q���ۤv���a�A�o�̥i�H�����z���U�حn�D�G(�{�b��X�Фl����o���Ӫ�1/5������)</font></marquee></font></p>
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#6633FF">
        <tr> 
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu.asp><font color="#FF3333">�@��Ы�</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu3.asp><font color="#FF3333">���Ť��J</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu4.asp><font color="#FF3333">���v��</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu5.asp><font color="#FF3333">���اO��</font></a></div>
          </td>
        </tr>
      </table>
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="15" width="14%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�Ы�����</font></div>
    </td>
    <td height="15" width="9%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">���</font></div>
    </td>
    <td height="15" width="11%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">����</font></div>
    </td>
    <td height="15" colspan="2" bgcolor="#FF9933" bordercolor="#FF6600">
      <div align="center"><font size="2" color="#FF3333">���P���X</font></div>
    </td>
    <td height="15" width="19%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">�ХD/��Q</font></div>
    </td>
    <td height="15" width="25%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�ާ@</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor=#DEAD63> 
    <td width="14%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("����")%> 
      </center>
      </font> 
    <td width="9%" bgcolor="#FFFFFF" class="calen-curr" height="15"><%=rs("���")%> 
      <div align="center"></div>
    <td width="11%" bgcolor="#FFFFFF" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("��m")%></font></div>
    </td>
    <td colspan="2" bgcolor="#FFFFFF" class="ct-def4" height="15"><%=rs("��D")%> 
      <font color=red><%=rs("id")%></font><font size="2">��</font></td>
    <td width="19%" bgcolor="#FFFFFF" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("��D")%> / </font><font color="ff00ff"><%=rs("��Q")%></font></div>
    </td>
    <td width="25%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("��D")<>name and rs("��D")="�L" then%><a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">�ʶR</font></a><%end if%><%if rs("��D")=name then%><a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>��X</font></a><%end if%><%if rs("��D")<>"�L" and rs("��D")<>name then%>�өФv�X��<%end if%> 
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 
conn.close 
%> 
</table>
</html>