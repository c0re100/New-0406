<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=Request.QueryString("ID")
if info(6)<>"�x��" then Response.Redirect "../error.asp?id=451"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
rs.open "select �P�� FROM �Τ� WHERE id=" & info(9),conn
tongmen=rs("�P��")
rs.close
rs.open "Select ����,²��,�x��,�A�X from ���� where ����='"&Id&"'",conn
if rs.eof or rs.bof then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('�L�������I�I');location.href = 'javascript:history.back()';</script>"
			response.end
end  if
%>
<html>
<head>
<title>�������e�ק�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#660000" text="#FFFFFF" link="#FFFF00" alink="#FFFF00" vlink="#FFFF00">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="498" height="31">
<tr align="center" bgcolor="#336633">
      <td width="100%" height="25" bgcolor="#99CCFF"><b><font color="#FF0000" size="4">�����޲z</font></b></td>
</tr>
</table>


</div>
<form action="updatmp.asp?subject=<%=rs("����")%>" method=POST>
  <ul>
    <table border=1 cellspacing=0 cellpadding=1 align="center" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
      <tr> 
        <td width="180"><font size="2" class="c" color="#FF6600">����</font></td>
        <td width="308"> <font color="#FF6600"><%=info(5)%> </font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">²��</font></td>
        <td width="308"> <font color="#FF6600"> 
          <input name="comment" value="<%=RS("²��")%>" size=40 maxlength=50>
          </font></td>
      </tr>
      <tr> 
        <td width="180"><font class="c" size="2" color="#FF6600">�x��&nbsp;</font></td>
        <td width="308"> <font color="#FF6600"><%=rs("�x��")%></font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">�A�X(�k/�k/�Ҧ�)</font></td>
        <td width="308"> <font color="#FF6600"> 
          <input name="shihe" value="<%=RS("�A�X")%>" size=40 maxlength=100>
          </font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">�x���I��]�����̤l�^</font></td>
        <td width="308"><font color="#FF6600"> 
          <input name="zhangmen" size=10 maxlength=10 value="�L">
          </font></td>
      </tr>
      <tr>
        <td width="180"><font color="#FF6600">�P������</font></td>
        <td width="308"><font color="#FF6600" size="2"><%=tongmen%></font></td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c">�@�ӤH�����@�Ӫ������x���A���ର�h�a���ڡA�H�̷s���]�m���ǡC</font><font size="3" class="c" color="#000000"><br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="��s" style="border: 2px solid;background-color:#FCCFDF; font-size: 9pt; border-color:
#000000 solid">
        </font></p>
    </div>
  </ul>
</form>
<%rs.close
rs.open "SELECT �m�W FROM �Τ� where ����='"&info(5)&"' and ����='�ƴx��'",conn
%>

<table width="100%" border="0">
  <tr>�ƴx���G
  <%
do while not rs.bof and not rs.eof
%>  
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("�m�W")%>' title="��������"><%=rs("�m�W")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT �m�W FROM �Τ� where ����='"&info(5)&"' and ����='����'",conn
%>
<table width="100%" border="0">
  <tr>���ѡG 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("�m�W")%>' title="��������"><%=rs("�m�W")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT �m�W FROM �Τ� where ����='"&info(5)&"' and ����='�@�k'",conn
%>
<table width="100%" border="0">
  <tr>�@�k�G 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("�m�W")%>' title="��������"><%=rs("�m�W")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT �m�W FROM �Τ� where ����='"&info(5)&"' and ����='��D'",conn
%>
<table width="100%" border="0">
  <tr>��D�G 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("�m�W")%>' title="��������"><%=rs("�m�W")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<div align="center">
  <p>&nbsp;</p>
  <p>&nbsp;<input type=button value=�������f onClick='window.close()' name="button" style="border: 2px solid;background-color:#FCCFDF; font-size: 9pt; border-color:
#000000 solid"></p>
</div>
</body>
</html>
<%
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
%> 