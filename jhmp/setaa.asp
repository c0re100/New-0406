<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
meme=info(0)
if info(4)=0 then
	Response.Write "<script language=JavaScript>{alert('�D�|�������}�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
%>
<html>
<head>
<title>�Z�\�]�m</title>
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
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: �s�ө���; font-size: 9pt; font-style: bold;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body background="../ff.jpg">
<table cellpadding="1" cellspacing="0" border=0 align="center" width="597" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr bgcolor="#003366"> 
<td width=10%><div align="center"><font color="#FF6600">ID</font></div></td>
<td width=15%><div align="center"><font color="#FF6600">���[��</font></div></td>
<td width=25%><div align="center"><font color="#FF6600">����</font></div></td>
<td width=50%><div align="center"><font color="#FF6600">�I�k����</font></div></td>
</tr>
<%
rs.open "SELECT * FROM �ϥΧޯ� where �m�W='"&info(0)&"'",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
<form method=POST action='setaa1.asp?a=m&id=<%=rs("id")%>'>
<tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
<td><div align="center"><font color="#FF6600"><%=rs("id")%></font></div></td>
<td><div align="center"><input type="text" name="wg1" size="12" value="<%=rs("�ޯ�")%>" maxlength="10"></div></td>
<td><div align="center"><input type="text" name="wg2" size="20" value="<%=rs("����")%>"></div></td>
<td><div align="center"><input type="text" name="wg3" size="50" value="<%=rs("�I�k����")%>"></div></td>
</tr>
<tr><td colspan=4 bgcolor=ffcc00> 
�Τ�G <input type="text" name="name" size="10" value=<%=meme%>>
�K�X�G <input type="password" name="pass" size="10" maxlength="10">
<input type="submit" value="�T�w�ק�" name="submit"></td></tr>
</form>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
