<HTML><HEAD><TITLE>����楫��</TITLE><META content="text/html; charset=big5" http-equiv=Content-Type><LINK
href="../chat/readonly/style.css" rel=stylesheet type=text/css>
<META content="MSHTML 5.00.2614.3500" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff text=#000000 topMargin=0 background="../ljimage/11.jpg">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=608><TBODY> <TR><TD height=158 vAlign=top width="608"><TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain
width=560> <TBODY><TR><TD vAlign=top><BR><TABLE align=center border=0
cellPadding=0 cellSpacing=0 class=table1 width="100%"> <TBODY><TR><TD align=middle vAlign=center><P class=p9><font color="#FF0000"><b>�i���򴣥ܡj</b></font><br>�o��O���򤤳̤j���楫���A�A�i�H�b�o�R�찵���Ϊ��D�ơA���ƩM���ơA�R�F��ƴN���h���򥩭��]���@�������������@�C<br></P><table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010"><tr bgcolor="#FFFFFF"><td height="16" colspan="5"><p align=center class="p9"><font color="#FF3333"><b>����楫��</b></font></p><tr bgcolor=#74E76D bordercolor="#000000"><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="12%"><div align="center"><font size="2" color="#FF0000">��W</font></div></td><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="18%"><div align="center"><font size="2" color="#FF0000">����</font></div></td><td height="15" width="16%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">�W�[������</font></div></td><td height="15" width="22%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">���</font></div></td><td height="15" width="32%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">�ʶR</font></div></td></tr><%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id,���~�W,����,������,�Ȩ� FROM ��� "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%> <tr bgcolor=#DEAD63><td bgcolor="#990000"><center><font size="2" color="#FFFFFF"><%=rs("���~�W")%>
</font></center><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("����")%></font></div><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("������")%></font></div></td><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("�Ȩ�")%></font></div></td><td bgcolor="#990000"><center><font size="2"><a href=caichang1.asp?id=<%=rs("id")%>><span class="calen-curr"><font color="#00FF00">�ʶR</font></span></a></font></center></td></tr><%
rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
%> </table>
                  <P align=center class="p9"><b><a href="qmg.htm"><font color="#0000FF">�i��^����j</font></a></b></P>
                  <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE><P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR><TR><TD bgColor=#847939 height=1 width="608"><IMG height=1 src="../wupin/jh.files/page_point.gif"
width=1></TD></TR></TBODY></TABLE>
</BODY></HTML> 