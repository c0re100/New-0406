<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
action=request.querystring("action")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���~�W,�ƶq,��O,���O,�Ȩ� FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>0 and ����='�ī~'  order by �Ȩ� ",conn
%>
<html>
<head>
<title>�ڪ��]��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body background="back.gif">
<div align="left"> <table width="453" border="0" cellspacing="0" cellpadding="0" height="224" align="center"> 
<tr> <td width="92" rowspan="3" valign="top"> <div align="center"><br> <input onClick="javascript:window.document.location.href='index.asp'" title=�˳Ƥ@�� type=button value="�˳Ƥ@��" name="button7"> 
<br> <br> <input onClick="javascript:window.document.location.href='head.asp'" title=�Z���˳� type=button value="�Z���˳�" name="button"> 
<br> <br> <input onClick="javascript:window.document.location.href='eat.asp'" title=����Ī� type=button value="����Ī�" name="button62"> 
</div></td><td valign="top" rowspan="3" width="361"> <div align="center"><img src="dao.gif" width="40" height="15">����ī~�@��<img src="dao1.gif" width="40" height="15"> 
<font color="#CC0000" face="����"><a href="javascript:this.location.reload()">��s</a></font></div><div align="center"> 
�`�N:��ܪA�ΡA�N�A�ΩҦ����ī~�I<br> <br> <table border="1" align="center" width="326" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000"> 
<tr align="center"> <td nowrap width="65" height="15"> <div align="center"><font color="#000000">���~�W</font></div></td><td nowrap width="30" height="15"> 
<div align="center"><font color="#000000">�ƶq </font> </div><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">�[��O</font></div></td><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">�[���O</font></div></td><td colspan="2" nowrap height="15"> 
<div align="center"><font color="#000000">����</font></div></td><td width="49" height="15"> 
<div align="center"><font color="#000000">�ϥμƶq</font></div></td><td width="63" height="15"> 
<div align="center"><font color="#000000">�ާ@</font></div></td></tr> <%
do while not rs.eof
%> <tr> <form method=POST action='wupin1.asp?action=<%=rs("���~�W")%>&name=<%=info(0)%>'> 
<td width="65" height="10"><div align="center"><%=rs("���~�W")%> </div></td><td width="30" height="10"> 
<div align="center"><font color="#FF0000"><%=rs("�ƶq")%> </font> </div><td width="36" height="10"> 
<div align="center"><font color="#0000FF"><%=rs("��O")%></font> </div></td><td width="36" height="10"> 
<div align="center"><font color="#0000FF"><%=rs("���O")%> </font> </div></td><td colspan="2" height="10"> 
<div align="center"><font color="#FF0000"><%=rs("�Ȩ�")%> </font> </div></td><td width="49" height="10"> 
<div align="center"> <select name="wpsl"> <option value="1" selected>1</option> 
<option value="2">2</option> <option value="3">3</option> <option value="4">4</option> 
<option value="10">10</option> <option value="15">15</option> <option value="20">20</option> 
<option value="25">25</option> <option value="50">50</option> </select> </div></td><td width="63" height="10"> 
<div align="center"> <input type="SUBMIT" name="Submit2"  value="�i��"> </div></td></form></tr> 
<%
rs.movenext
loop
%> </table><br> <%
rs.close
rs.open "SELECT ���~�W,�ƶq,�\��,�W�[�ƭ� FROM �׽m���~ WHERE �֦���=" & info(9) & " and �ƶq>0  order by �W�[�ƭ� ",conn
%> �׽m�ī~�C��<br> <table border="1" align="center" width="326" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000"> 
<tr align="center"> <td nowrap width="65" height="15"> <div align="center"><font color="#000000">���~�W</font></div></td><td nowrap width="30" height="15"> 
<div align="center"><font color="#000000">�ƶq </font> </div><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">�\��</font></div></td><td colspan="2" nowrap height="15"> 
<div align="center"><font color="#000000">����</font></div></td><td width="49" height="15"> 
<div align="center"><font color="#000000">�ϥμƶq</font></div></td><td width="63" height="15"> 
<div align="center"><font color="#000000">�ާ@</font></div></td></tr> <%
do while not rs.eof
%> <tr> <form method=POST action='xleat.asp?action=<%=rs("���~�W")%>&name=<%=name%>'> 
<td width="65" height="17"><div align="center"><%=rs("���~�W")%> </div></td><td width="30" height="17"> 
<div align="center"><font color="#FF0000"><%=rs("�ƶq")%> </font> </div><td width="36" height="17"> 
<div align="center"><font color="#0000FF"><%=rs("�\��")%> </font> </div></td><td colspan="2" height="17"> 
<div align="center"><font color="#FF0000"><%=rs("�W�[�ƭ�")%> </font> </div></td><td width="49" height="17"> 
<div align="center"> <select name="xlsl"> <option value="1" selected>1</option> 
<option value="2">2</option> <option value="3">3</option> <option value="4">4</option> 
<option value="5">5</option> <option value="6">6</option> <option value="7">7</option> 
<option value="8">8</option> <option value="9">9</option> </select> </div></td><td width="63" height="17"> 
<div align="center"> <input type="SUBMIT" name="Submit22"  value="�i��"> </div></td></form></tr> 
<%
rs.movenext
loop

%> </table><BR><%
rs.close
rs.open "SELECT ���~�W,�ƶq,���O,�Ȩ� FROM ���~ WHERE �֦���=" & info(9) & " and ����='�A��' and �ƶq>0  order by �Ȩ� ",conn
%> �A��C��<BR><TABLE BORDER="1" ALIGN="center" WIDTH="326" CELLPADDING="0" CELLSPACING="0" HEIGHT="23" BORDERCOLOR="#000000"> 
<TR ALIGN="center"> <TD nowrap WIDTH="65" HEIGHT="15"> <DIV ALIGN="center"><FONT COLOR="#000000">���~�W</FONT></DIV></TD><TD nowrap WIDTH="30" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">�ƶq </FONT> </DIV><TD nowrap WIDTH="36" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">�\��</FONT></DIV></TD><TD COLSPAN="2" nowrap HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">����</FONT></DIV></TD><TD WIDTH="49" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">�ϥμƶq</FONT></DIV></TD><TD WIDTH="63" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">�ާ@</FONT></DIV></TD></TR> <%
do while not rs.eof
%> <TR> <FORM METHOD=POST ACTION='eathua.asp?action=<%=rs("���~�W")%>&name=<%=name%>'> 
<TD WIDTH="65" HEIGHT="17"><DIV ALIGN="center"><%=rs("���~�W")%> </DIV></TD><TD WIDTH="30" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#FF0000"><%=rs("�ƶq")%> </FONT> </DIV><TD WIDTH="36" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#0000FF"><%=rs("���O")%> </FONT> </DIV></TD><TD COLSPAN="2" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#FF0000"><%=rs("�Ȩ�")%> </FONT> </DIV></TD><TD WIDTH="49" HEIGHT="17"> 
<DIV ALIGN="center"> <SELECT NAME="hua"> <OPTION VALUE="1" selected>1</OPTION> 
<OPTION VALUE="2">2</OPTION> <OPTION VALUE="3">3</OPTION> <OPTION VALUE="4">4</OPTION> 
<OPTION VALUE="5">5</OPTION> <OPTION VALUE="6">6</OPTION> <OPTION VALUE="7">7</OPTION> 
<OPTION VALUE="8">8</OPTION> <OPTION VALUE="9">9</OPTION> </SELECT> </DIV></TD><TD WIDTH="63" HEIGHT="17"> 
<DIV ALIGN="center"> <INPUT TYPE="SUBMIT" NAME="Submit222"  VALUE="�i��"> </DIV></TD></FORM></TR> 
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%> </TABLE><br> <input onClick="JavaScript:window.close()" title=�������f type=button value="�������f" name="button2"> 
</div></td></tr> <tr> </tr> <tr> </tr> </table></div>
</body>
</html>
