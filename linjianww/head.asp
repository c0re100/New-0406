<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>

<head>
<title>�L���Q</title>

<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body text="#000000" background="0064.jpg" link="#000000" vlink="#000000" alink="#000000">
<table border="1" cellspacing="1" cellpadding="0" align="center" width="80%" bordercolor="#000000">
  <tr> 
    <td height="9" width="14%"> 
      <div align="center"><a href="binqi.asp">����M�C</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"><a href="head.asp">�Y��</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"><a href="body.asp">����</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"><a href="foot.asp">���}</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"></div>
      <div align="center"><a href="other.asp">�˹�</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"><a href="anqi.asp">�t��</a></div>
    </td>
    <td height="19" width="14%" rowspan="2"> 
      <div align="center"><a href="addfy.asp">�K�[�˳�</a></div>
    </td>
  </tr>
  <tr> 
    <td height="9" width="14%"> 
      <div align="center"><a href="dunpai.asp">����@��</a></div>
    </td>
  </tr>
</table>
<p align="center">�Y���޲z</p>
<br>
<table border="1" align="center" width="600" cellpadding="0"
cellspacing="0">
  <tr bgcolor="#00CCFF"> 
    <td height="17" colspan="6"> 
      <p align="center">�{ �� �Y ��</p>
    
  <tr> 
    <td height="18" bgcolor="#C4DEFF" width="75"> 
      <div align="center"> <font size="2" color="#000000">���ҦW</font> </div>
    </td>
    <td bgcolor="#C4DEFF" height="18" width="92"> 
      <div align="center"><font size="2" color="#000000">�� ��</font> </div>
    </td>
    <td bgcolor="#C4DEFF" height="18" width="148"> 
      <div align="center"><font size="2" color="#000000">�Ϥ�����</font></div>
    </td>
    <td bgcolor="#C4DEFF" height="18" width="132"> 
      <div align="center"><font size="2" color="#000000">�\ ��</font></div>
    </td>
    <td height="18" bgcolor="#C4DEFF" width="65"> 
      <div align="center"> <font size="2" color="#000000">�� ��</font> </div>
    </td>
    <td height="18" bgcolor="#C4DEFF" width="74"> 
      <div align="center"> <font size="2" color="#000000">�� �@</font> </div>
    </td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM ���~�R where  ����='�Y��' order by �Ȩ�"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td height="11" width="75"><%=rs("���~�W")%> 
    <td height="11" width="92"> 
      <div align="center"><img src="../hcjs/jhjs/001/<%=rs("����")%>.gif" border="0" alt="<%=rs("���~�W")%> "></div>
    </td>
    <td height="11" width="148">../hcjs/jhjs/001/<%=rs("����")%>.gif</td>
    <td height="11" width="132">�W����<%=abs(rs("����"))%> �W���s<%=abs(rs("���s"))%></td>
    <td height="11" width="65"><%=rs("�Ȩ�")%>��</td>
    <td height="11" width="74"> 
      <div align="center"><a href="modifybq.asp?wupinid=<%=rs("id")%>">�ק�</a> 
        <a href="del.asp?id=<%=rs("id")%>">�R��</a></div>
    </td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
</body>

</html>
