<%@ LANGUAGE=VBScript codepage ="950" %>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ljjmwb.asa")
Conn.Open connstr
rs.open "SELECT ���a�I�� FROM mess",conn,2,2
%>
<html>
<head>
<title>�F�C���a�[��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body bgcolor="#9933CC" text="#000000" background="../2.jpg">
<div align="center">
  <p><b><font color="#0000CC" size="+2">�F�C������a�p��<br>
    </font></b></p>
  <p align="left">1.���a�[��������H<br>
    a.�㦳�@�w�W�Ҫ����a�C<br>
    <br>
    b.�i�H�����u��W���A��.<br>
    <br>
    c.�b���a�Ϩ�Ƥ@�w�v�T�O�����a.<br>
    <br>
    d.�ݦV�F�C���򴣨Ѻ��a�`�U��ơI<br>
    <br>
    e.���IP�a�}���T�w�̡A�ݦۦ�ק�IP�]�m�C<br>
    <br>
    2.�[�J���O�зǤ��u�f�����H<br>
    <br>
    a.�[�����a���O�G200��/��(��B���B�O�Ψ䥦�a�ϥH��������200$��/��!)<br>
    (<b><font color="#FF0000">�K�`�����S��:150/��A��B���B�O���ܡI</font></b>) <br>
    <br>
    b.�b�[�����a�W���C�ѥi�H�h����u��1000�U��A�|�O10���B����10��,�y�P�@�ӡI<br>
    <br>
    c.�b�[�����a�W���A�_���᪬�A���|�ᥢ�I<br>
    <br>
    d.�b�[�����a�W���A<b><font color="#0000FF">�w�I�O���`���a�W����2���I </font></b><br>
    <br>
    e.�F�C����N�|�b�����u�ʼs�i����@�P�I<br>
    <br>
    f.�b�[�����a�W���ʪ���7��I(�|������)<br>
    <br>
    <br>
    <br>
    3.�[����k�G<br>
  </p>
  <table width="75%" border="1" align="left">
    <tr>
      <td bgcolor="#9999CC"><%=rs("���a�I��")%> 
        <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
      </td>
    </tr>
  </table>
  <p align="left"> <br>
  </p>
</div>
</body>
</html>
