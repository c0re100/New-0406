<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<%info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
my=info(0)
%>
<html>
<head>
<title>�d���ө���</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" leftmargin="0" topmargin="7">
<div align="left"></div>
<div align="center">
  <table width="640" border="0" cellspacing="0" cellpadding="4">
    <tr> 
      <td width="615" colspan="2" align="center" valign="top"> [ <font color="#009900">�d �� 
        �� ��</font> ]<br> <br>
        �i<a href="indexsheep.asp"> �d���ө� </a>�j �i<a href="stunt.asp"> �ޯ�ө�</a> �j 
        �i <a href="itemshop.asp">�D��ө� </a>�j �i<a href="sellsheep.asp"> <font color="#FF6600">�ڭn�汼�ڪ��d��</font></a> 
        �j <br> <br> <div align="center"> </div>
        <div align="center"><font size="-1">�w����{�d���ө��A�o�̦��U�����P�������d�ѧA���ʨ��C�`�N�A�d������R�@�����C</font><br>
          <!--#include file="data.asp"-->
          <table width="620" border="1" cellspacing="2" cellpadding="0" bordercolor="#FF0000" height="26">
            <tr bgcolor="#FF0000"> 
              <td width="145" height="17"> <div align="center"><font color="#FFFFFF">�d������ 
                  </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">�d���ޯ�</font></div></td>
              <td width="187" height="17"> <div align="center"><font color="#FFFFFF">�d���Ѽ� 
                  </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">�� 
                  �� </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">�� 
                  �@ </font></div></td>
            </tr>
            <%
sql="SELECT �d������,�ޯ�,����,����,���s,��O,���� FROM �d��'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
            <tr> 
              <td width="145" height="22"><img src="image/<%=rs("����")%>.gif" border="0" alt="<%=rs("�d������")%> " align="absmiddle"></td>
              <td width="72" height="22"> <div align="center"><%=rs("�ޯ�")%></div></td>
              <td width="187" height="22">�����G<%=rs("����")%> ���s�G<%=rs("���s")%> ��O�G<%=rs("��O")%> </td>
              <td width="72" height="22"> <div align="center"><%=rs("����")%>��</div></td>
              <td width="72" height="22"> <p align="center"><a href="buy.asp?name=<%=rs("�d������")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></p></td>
            </tr>
            <%
rs.movenext
loop
conn.close
%>
          </table>
          <br>
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>