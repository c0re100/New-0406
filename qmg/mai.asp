<HTML><HEAD><TITLE>�F�C����-��ѥ����~��</TITLE><META content="text/html; charset=big5" http-equiv=Content-Type><LINK
href="../chat/readonly/style.css" rel=stylesheet type=text/css>
<META content="MSHTML 5.00.2614.3500" name=GENERATOR></HEAD><BODY bgColor=#ffffff text=#000000 topMargin=0><TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=608><TBODY><TR><TD  vAlign=top width=608><TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9
width="100%"> <TBODY><TR><TD height=25 width="31%">&nbsp;</TD><TD height=25 vAlign=top width="33%"><DIV align=center><font color="#000000">�F�C���򥩭��~��</font></DIV></TD><TD height=25 vAlign=top
width="36%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR> <TR><TD height=158 vAlign=top width="608"><TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain
width=560> <TBODY><TR><TD vAlign=top><BR><TABLE align=center border=0
cellPadding=0 cellSpacing=0 class=table1 width="100%"> <TBODY><TR><TD align=middle vAlign=center><P class=p9><font color="#FF0000"><b>�i���򴣥ܡj</b></font><br>�o��O���򤤳̤j�������]�~�泡�A�A�i�H�b�o��R�O�H�������A�ӼW�[��O�@~�I�I�I<br></P><table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010"><tr bgcolor="#FFFFFF"><td height="16" colspan="5"><p align=center class="p9"><font color="#FF3333"><b>���򥩭��~��</b></font></p><tr bgcolor=#74E76D bordercolor="#000000"><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="12%"><div align="center"><font size="2" color="#FF0000">���W</font></div></td><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="18%"><div align="center"><font size="2" color="#FF0000">�����v��</font></div></td><td height="15" width="16%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">�W�[��O</font></div></td><td height="15" width="22%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">���</font></div></td><td height="15" width="32%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">��</font></div></td></tr><%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �歱 where �ɭ�=false"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%> <tr bgcolor=#DEAD63><td bgcolor="#990000"><center><font size="2" color="#FFFFFF"><%=rs("id")%>
</font></center><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("�֦���")%></font></div><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("��O")%></font></div></td><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("���")%></font></div></td><td bgcolor="#990000"><center><font size="2"><a href=chi1.asp?id=<%=rs("id")%>><span class="calen-curr"><font color="#00FF00">��</font></span></a></font></center></td></tr><%
rs.movenext
loop
set rs=nothing
conn.close%> </table><P align=center class="p9"><b><a href="qmg.htm"><font color="#74E76D">�i��^���]�j</font></a></b></P><P align=center>&nbsp;</P></TD></TR></TBODY></TABLE><P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR><TR><TD bgColor=#847939 height=1 width="608"><IMG height=1 src="../wupin/jh.files/page_point.gif"
width=1></TD></TR></TBODY></TABLE>
</BODY></HTML> 