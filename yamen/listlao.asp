<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>�b��H��</title>
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
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#006666" background="../ff.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0">
<br>
<table border=0 cellspacing="2" cellpadding="2" width="560" align="center">
  <tr bgcolor="#CC0000"> 
    <td width="15%"> 
      <div align="center"><font color="#FFFFFF"> �ǤH </font></div></td>
    <td width=15%> 
      <div align="center"><font color="#FFFFFF"> ���� </font></div></td>
    <td width=11%> 
      <div align="center"><font color="#FFFFFF"> ���� </font></div></td>
    <td width=27%> 
      <div align="center"><font color="#FFFFFF"> �� �A </font></div></td>
    <td colspan=2> 
      <div align="center"><font color="#FFFFFF">�� �@ </font></div></td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �m�W,����,����,���A,id,�n�� FROM �Τ� where ���A='��' or ���A='�c'",conn
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#330000"> 
    <td><font color="#FFFFCC"><%=rs("�m�W")%></td>
    <td><font color="#FFFFCC"><%=rs("����")%></td>
    <td><%=rs("����")%></td>
    <%if rs("���A")="��" then %>
    <td width="27%"><div align="center"> <a href="shifang.asp?name=<%=rs("�m�W")%>">���100�U������</a>|<a href="jieyu.asp?name=<%=rs("�m�W")%>" title="���\�v�u��30%">�T��</a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>">����</a> 
        <%end if%>
      </div></td>
    <% else
if rs("�n��")>now() then%>
    <td width="23%" height="31"> <div align="center"> <font color="#FFFFCC">�L�v�ާ@</font> 
        <%if info(0)="�����`��" then%>
        <a href="shifang.asp?name=<%=rs("�m�W")%>">���50000������</a> 
        <%else%>
        �L�v�ާ@ 
        <%end if%>
        ����O��|����ɶ��G<%=rs("�n��")%><a href="yongka.asp?name=<%=rs("�m�W")%>">�Υd</a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>">����</a> 
        <%end if%>
        |<a href="jieyu.asp?name=<%=rs("�m�W")%>"
title="���\�v�u��30%">�T��</a></div></td>
    <%else%>
    <td width="9%"> <div align="center">�w����</div></td>
    <% conn.execute("update �Τ� set ���A='���`',�n��=now() where �m�W='"&rs("�m�W")&"'")
end if%>
    <%end if%>
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
<p align="center"><a href="#" onClick="window.close()">�� ��</a></p>
</body>
</html>
