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
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background=../bk.jpg link="#0000FF" vlink="#0000FF" alink="#0000FF">
<br>
<table border="1" cellspacing="0" cellpadding="0" width="560" align="center">
  <tr> 
    <td width="15%" nowrap bgcolor="#99CCFF"> 
      <div align="center"> �ǤH </div>
    <td width="15%" nowrap bgcolor="#99CCFF"> 
      <div align="center"> ���� </div>
    </td>
    <td width="11%" bgcolor="#99CCFF"> 
      <div align="center"> ���� </div>
    </td>
    <td width="27%" bgcolor="#99CCFF"> 
      <div align="center"> �� �A</div>
    </td>
    <td colspan="2" bgcolor="#99CCFF"> 
      <div align="center">�� �@ </div>
    </td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,�m�W,����,����,���A,�n�� FROM �Τ� where ���A='��' or ���A='�c'",conn
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#FFCC99"  onMouseOut="this.bgColor='#FFCC99';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td width="15%" height="31"><%=rs("�m�W")%> 
    <td width="15%" height="31"><%=rs("����")%></td>
    <td width="11%" nowrap height="31"><%=rs("����")%></td>
    <%if rs("���A")="��" then %>
    <td width="27%" height="31"> 
      <div align="center"> <a href="shifang.asp?name=<%=rs("�m�W")%>"><font color="#0000FF">���100�U������</font></a>|<a href="jieyu.asp?name=<%=rs("�m�W")%>"
title="���\�v�u��30%"><font color="#FF0000">�T��</font></a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>"><font color="#FF0000">����</font></a> 
        <%end if%>
      </div>
    </td>
    <% else
if rs("�n��")>now() then%>
    <td width="23%" height="31"> 
      <div align="center"> <font color="#0000FF">�L�v�ާ@</font> 
        <%if  info(0)="�����" then%>
        <a href="shifang.asp?name=<%=rs("�m�W")%>"><font color="#0000FF">���50000������</font></a> 
        <%else%>
        <font color="#0000FF">�L�v�ާ@</font> 
        <%end if%>
        <font color="#FF0033">����O��</font>|����ɶ��G<%=rs("�n��")%><a href="yongka.asp?name=<%=rs("�m�W")%>"><font color="#0000FF">�Υd</font></a> 
        <%if info(2)>=6 then%>
        |<a href="shuchu.asp?id=<%=rs("id")%>"><font color="#FF0000">����</font></a> 
        <%end if%>
        |<a href="jieyu.asp?name=<%=rs("�m�W")%>"
title="���\�v�u��30%"><font color="#FF0000">�T��</font></a> </div>
    </td>
    <%else%>
    <td width="9%" height="31"> 
      <div align="center"> <font color="#FF0033">�w����</font> </div>
    </td>
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

<html></html> 



