<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
toname=Trim(Request.QueryString("toname"))
if toid=0 then toid=info(9)
if toname="�j�a" then toname=info(0)
if info(2)<10 then
Response.Write "<script Language=Javascript>alert('�A�Q�@����u�I');window.close();</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM �Τ� WHERE �m�W='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "SELECT * FROM ���~ WHERE �֦���=" & idd & " and �ƶq>0  order by ���� ",conn
%>
<html>
<head>
<title>���~�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: �s�ө���; font-size: 9pt; font-style: bold;}
</style>
<style TYPE="text/css">
body{overflow:scroll;overflow-x:hidden}
.popper{  position : absolute;  visibility : hidden;}
</style>
<STYLE TYPE="text/css">
BODY {OVERFLOW:scroll;OVERFLOW-X:hidden}
.DEK {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}
</STYLE>
</head>
<body bgcolor="#660000"  background=../ff.jpg leftmargin="0" topmargin="0">
<div align="center"><%=toname%>�{�����~�@��(�˳ƪ����~���~)<a href="javascript:this.location.reload()">��s</a></div>
<br>
<table border="1" align="center" width=80% cellpadding="1" cellspacing="0" height="31">
<tr align=center>
<td>���~�W</td><td>����</td><td>�ƶq</td><td>���O</td><td>��O</td><td>����</td><td>���s</td><td>����</td><td>��T��</td><td>����</td><td>�˳�</td><td>�覡</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td><%=rs("���~�W")%> </div></td>
<td><%=rs("����")%></div></td>
<td><%=rs("�ƶq")%></div></td>
<td><%=rs("���O")%></div></td>
<td><%=rs("��O")%></div></td>
<td><%=rs("����")%></div></td>
<td><%=rs("���s")%></div></td>
<td><%=rs("����")%></div></td>
<td><%=rs("��T��")%></div></td>
<td><%=rs("�Ȩ�")%></div></td>
<td><%=rs("�˳�")%></div></td>
<td><a href="del.asp?id=<%=rs("id")%>">�R��</a><br><a href="decc.asp?id=<%=rs("id")%>">�ƶq�[10</a><br><a href="decc20.asp?id=<%=rs("id")%>">�ƶq�[20</a><br><a href="decc100.asp?id=<%=rs("id")%>">�ƶq�[100</a></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM ������� WHERE �֦���=" & idd & " order by �覡",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="45" height="16"><div align="center"> ���~�W </div></td>
<td nowrap width="40" height="16"><div align="center"> ����  </div></td>
<td nowrap width="33" height="16"><div align="center"> �ƶq   </div></td>
<td nowrap width="40" height="16"><div align="center"> ���O  </div></td>
<td nowrap width="40" height="16"><div align="center"> ��O  </div></td>
<td nowrap width="40" height="16"><div align="center"> ����  </div></td>
<td nowrap width="40" height="16"><div align="center"> ���s  </div></td>
<td colspan="2" nowrap height="16"><div align="center"> ���� </div></td>
<td nowrap width="45" height="16"><font size="-1">�覡 </td>
<td nowrap width="50" height="16"><div align="center"> �ާ@ </div></td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td width="45" height="3"><div align="center"> <%=rs("���~�W")%>  </div></td>
<td width="40" height="3"><div align="center"> <%=rs("����")%> </div></td>
<td width="33" height="3"><div align="center"> <%=rs("�ƶq")%>  </div></td>
<td width="40" height="3"><div align="center"> <%=rs("���O")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("��O")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("����")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("���s")%> </div></td>
<td colspan="2" height="3"><div align="center"> <%=rs("�Ȩ�")%>  </div></td>
<td height="3" width="45"><div align="center"> <%=rs("�覡")%> </div></td>
<td height="3" width="50"><div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">�R�� </a> </div></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<tr align="center">
<td nowrap width="45" height="16">
<div align="center"></div>
</td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
<tr>
<td height="25">
<div align="center">�o�̥i�H�d�ݨ��誺���~�A�R���N�i�H�N���Τ᪺���~�R���I</div>
</td>
</tr>
</table>
</body>
</html>
