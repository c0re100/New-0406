<HTML><HEAD><TITLE>江湖菜市場</TITLE><META content="text/html; charset=big5" http-equiv=Content-Type><LINK
href="../chat/readonly/style.css" rel=stylesheet type=text/css>
<META content="MSHTML 5.00.2614.3500" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff text=#000000 topMargin=0 background="../ljimage/11.jpg">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=608><TBODY> <TR><TD height=158 vAlign=top width="608"><TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain
width=560> <TBODY><TR><TD vAlign=top><BR><TABLE align=center border=0
cellPadding=0 cellSpacing=0 class=table1 width="100%"> <TBODY><TR><TD align=middle vAlign=center><P class=p9><font color="#FF0000"><b>【江湖提示】</b></font><br>這兒是江湖中最大的菜市場，你可以在這買到做面用的主料，輔料和湯料，買了原料就塊去江湖巧面館做一份美味的巧面哦。<br></P><table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010"><tr bgcolor="#FFFFFF"><td height="16" colspan="5"><p align=center class="p9"><font color="#FF3333"><b>江湖菜市場</b></font></p><tr bgcolor=#74E76D bordercolor="#000000"><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="12%"><div align="center"><font size="2" color="#FF0000">菜名</font></div></td><td height="15" bgcolor="#66FF66" bordercolor="#FF6600" width="18%"><div align="center"><font size="2" color="#FF0000">類型</font></div></td><td height="15" width="16%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">增加美味度</font></div></td><td height="15" width="22%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">售價</font></div></td><td height="15" width="32%" bgcolor="#66FF66" bordercolor="#FF6600"><div align="center"><font size="2" color="#FF0000">購買</font></div></td></tr><%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id,物品名,類型,美味度,銀兩 FROM 菜場 "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%> <tr bgcolor=#DEAD63><td bgcolor="#990000"><center><font size="2" color="#FFFFFF"><%=rs("物品名")%>
</font></center><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("類型")%></font></div><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("美味度")%></font></div></td><td bgcolor="#990000"><div align="center"><font size="2" color="#FFFFFF"><%=rs("銀兩")%></font></div></td><td bgcolor="#990000"><center><font size="2"><a href=caichang1.asp?id=<%=rs("id")%>><span class="calen-curr"><font color="#00FF00">購買</font></span></a></font></center></td></tr><%
rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
%> </table>
                  <P align=center class="p9"><b><a href="qmg.htm"><font color="#0000FF">【返回江湖】</font></a></b></P>
                  <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE><P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR><TR><TD bgColor=#847939 height=1 width="608"><IMG height=1 src="../wupin/jh.files/page_point.gif"
width=1></TD></TR></TBODY></TABLE>
</BODY></HTML> 