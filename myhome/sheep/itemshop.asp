<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
%>
<html>
<head>
<title>�D��ө�</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" leftmargin="0" topmargin="7">
<div align="left"></div>
<div align="center">
  <table width="331" border="0" cellspacing="0" cellpadding="4">
    <tr> 
      <td colspan="2" valign="top" align="center"> [ <font color="#009900">�d �� 
        �� ��</font> ]<br> <br>
        �i<a href="indexsheep.asp"> �d���ө� </a>�j �i<a href="stunt.asp"> �ޯ�ө�</a> �j 
        �i <a href="itemshop.asp">�D��ө� </a>�j �i <a href="sellitem1.asp"><font color="#FF6600">�ڭn�汼�ڪ��D��</font></a> 
        �j <br> <br> <div align="center"> </div>
        <div align="center"><font size="-1">�w����{�d���ө��A�o�̦��U�����P�������d�ѧA���ʨ��C�`�N�A�d������R�@�����C</font><br>
          <!--#include file="data.asp"-->
          <br>
          <br>
          <table border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="#FF0000" width="600">
            <tr bgcolor="#FF0000"> 
              <td> <div align="center"><font color="#FFFFFF">�d������ </font></div></td>
              <td> <div align="center"><font color="#FFFFFF">�d���Ѽ� </font></div></td>
              <td> <div align="center"><font color="#FFFFFF">�� �� </font></div></td>
              <td> <div align="center"><font color="#FFFFFF">�� �@ </font></div></td>
            </tr>
            <%
sql="SELECT �D��W�r,����,���s,��O,���� FROM �d���D��'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
            <tr> 
              <td><%=rs("�D��W�r")%></td>
              <td> <div align="center"><span style="letter-spacing: 1">+�����G<%=rs("����")%> +���s�G<%=rs("���s")%> +��O�G<%=rs("��O")%> </span></div></td>
              <td> <div align="center"><span style="letter-spacing: 1"><%=rs("����")%>��</span></div></td>
              <td> <p align="center"><span style="letter-spacing: 1"><a href="buyitem.asp?name=<%=rs("�D��W�r")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></span></p></td>
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
          <br>
        </div></td>
    </tr>
  </table>
</div>
</body>

</html>