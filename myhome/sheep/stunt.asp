<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>�d���ޯ�ө�</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" background="../../ff.jpg" leftmargin="0" topmargin="7">
<div align="left"></div>
<div align="center">
  <table width="480" border="0" cellspacing="0" cellpadding="4">
    <tr> 
      <td colspan="2" valign="top" align="center"> [ <font color="#009900">�d �� 
        �� ��</font> ]<br> <br>
        �i<a href="indexsheep.asp"> �d���ө� </a>�j �i<a href="stunt.asp"> �ޯ�ө�</a> �j 
        �i <a href="itemshop.asp">�D��ө� </a>�j�i<a href="sellsheep.asp"> <font color="#FF6600">�ڭn�汼�ڪ��d��</font></a> 
        �j <br> <br> <div align="center"> </div>
        <div align="center"><font size="-1">�w����{�d���ө��A�o�̦��U�����P�������d�ѧA���ʨ��C�`�N�A�d������R�@�����C</font><br>
          <!--#include file="data.asp"-->
          <br>
          <table border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="#FF0000" width="600">
            <tr bgcolor="#FF0000"> 
              <td width="109"> <div align="center"><font color="#FFFFFF">�ޯ�W�� </font></div></td>
              <td width="190"> <div align="center"><font color="#FFFFFF">�� �G </font></div></td>
              <td width="114"> <div align="center"><font color="#FFFFFF">�� �� </font></div></td>
              <td width="167"> <div align="center"><font color="#FFFFFF">�� �@ </font></div></td>
            </tr>
            <%
sql="SELECT �ޯ�W��,²��,���� FROM �d���ޯ�'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
            <tr> 
              <td><%=rs("�ޯ�W��")%></td>
              <td><%=rs("²��")%> </td>
              <td> <div align="center"><%=rs("����")%>��</div></td>
              <td> <p align="center"><a href="buystunt.asp?name=<%=rs("�ޯ�W��")%>&money=<%=rs("����")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></p></td>
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