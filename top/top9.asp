<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�n���Ʀ�TOP20</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
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
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
const MaxPerPage=30
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>

<body text="#000000" bgcolor=000000 topmargin="0">
<div align="center">
<%
dim sql
dim rs
dim filename
sql="select �m�W,�ʧO,����,����,allvalue,mvalue,times,������� from �Τ� where ����='" & info(5) &"'order by ������� desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
 
response.write "<p align='center'>�S���i�Ʀ檺�ﹳ </p>"
else
%>
  <table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
    <tr bgcolor="#990000"> 
      <td align="center" height="16" width="9%"><font color="#FFFFFF">�m�W</font></td>
      <td align="center" height="16" width="6%"><font color="#FFFFFF">�ʧO</font></td>
      <td align="center" height="16" width="16%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" width="22%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" width="11%"><font color="#FFFFFF">�`�n��</font></td>
      <td align="center" height="16" width="8%"><font color="#FFFFFF">��n��</font></td>
      <td align="center" height="16" width="8%"><font color="#FFFFFF">�n������</font></td>
      <td align="center" height="16" width="13%"><font color="#FFFFFF">�������</font></td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td align="center" width="9%"><font color="#FFFFFF"><%=rs("�m�W")%></font></td>
      <td align="center" width="6%"><font color="#CCFFCC"><%=rs("�ʧO")%></font></td>
      <td align="center" width="16%"><font color="#CCFFCC"><%=rs("����")%></font></td>
      <td align="center" width="22%"><font color="#CCFFCC"><%=rs("����")%></font></td>
      <td align="center" width="11%"><font color="#CCFFCC"><%=rs("allvalue")%></font></td>
      <td align="center" width="8%"><font color="#CCFFCC"><%=rs("mvalue")%></font></td>
      <td align="center" width="8%"><font color="#CCFFCC"><%=rs("times")%></font></td>
      <td align="center" width="13%"><font color="#FFDDCC"><%=rs("�������")%> 
        </font></td>
    </tr>
    <%
rs.movenext
filename=filename+1
if filename>9 then Exit Do
loop
end if %>
  </table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#336633" height="16">
<tr>
      <td width="100%" bgcolor="#0033CC"><font size="2" color="#ffffff">�̤l�V�����^�m�����10��Ʀ�</font></td>
</tr>
</table>
<%
filename=0
sql="select * from �Τ� where ����='" & info(5) &"'order by -������� desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
 
response.write "<p align='center'>�S���i�Ʀ檺�ﹳ </p>"
else
%>
  <table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
    <tr bgcolor="#990000"> 
      <td align="center" height="16" width="9%"><font color="#FFFFFF">�m�W</font></td>
      <td align="center" height="16" width="6%"><font color="#FFFFFF">�ʧO</font></td>
      <td align="center" height="16" width="16%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" width="22%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" width="11%"><font color="#FFFFFF">�`�n��</font></td>
      <td align="center" height="16" width="8%"><font color="#FFFFFF">��n��</font></td>
      <td align="center" height="16" width="8%"><font color="#FFFFFF">�n������</font></td>
      <td align="center" height="16" width="13%"><font color="#FFFFFF">�������</font></td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td align="center" width="9%"><font color="#FFFFFF"><%=rs("�m�W")%></font></td>
      <td align="center" width="6%"><font color="#CCFFCC"><%=rs("�ʧO")%></font></td>
      <td align="center" width="16%"><font color="#CCFFCC"><%=rs("����")%></font></td>
      <td align="center" width="22%"><font color="#CCFFCC"><%=rs("����")%></font></td>
      <td align="center" width="11%"><font color="#CCFFCC"><%=rs("allvalue")%></font></td>
      <td align="center" width="8%"><font color="#CCFFCC"><%=rs("mvalue")%></font></td>
      <td align="center" width="8%"><font color="#CCFFCC"><%=rs("times")%></font></td>
      <td align="center" width="13%"><font color="#FFDDCC"><%=rs("�������")%> 
        </font></td>
    </tr>
    <%
rs.movenext
filename=filename+1
if filename>9 then Exit Do
loop
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 
%></table>
</div>
</body>
</html>
 
