<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


%>
<html>

<head>
<title>��}�Z��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#660000" background="../../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#000000" align="center" width="389" cellpadding="8"
cellspacing="8">
  <tr bgcolor="#FFFFFF"> 
    <td height="99" bgcolor="#FFFF00"> 
      <table border="1"
      width="409" cellspacing="2" cellpadding="1" bordercolor="#65251C">
        <tr> 
          <td colspan="5" height="33" bgcolor="#000000"> 
            <div align="center"> 
              <p><font size="+2" color="#FFFF00">��}�Z��</font></p>
              <p><font size="2" color="#FFFFCC">����G�s�]<img src="../../hcjs/jhjs/001/pig.gif" width="64" height="64" border="0" alt="�I�����_">�i�Ψӧ�}�Z��<br>
                ���ŶV����}���\���v�V�p�I</font><font size="+2" color="#FFFFCC"> </font></p>
            </div>
          </td>
        </tr>
      </table>
      <table width="355" border="1" align="center" cellspacing="1" cellpadding="1" bordercolor="#FF0000">
        <tr> 
          <td width="121"> 
            <div align="center"> 
              <%rs.open "SELECT id,���~�W,����,����,�ƶq,����,���s FROM ���~ WHERE �֦���=" & info(9) &" and �ƶq>0 and �˳�=false and (����='�Y��' or ����='����' or ����='�k��' or ����='���}' or ����='�˹�' or ����='����')  order by �Ȩ� ",conn%>
              <table width="395" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
                <tr bgcolor="#FF0000"> 
                  <td width="16%"> <div align="center"><font color="#FFFFFF">���~�W</font></div></td>
                  <td width="8%"> <div align="center"><font color="#FFFFFF">����</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">����</font></div></td>
                  <td width="10%"> <div align="center"><font color="#FFFFFF">�ƶq</font></div></td>
                  <td width="10%"> <div align="center"><font color="#FFFFFF">�[����</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">�[���s</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">�ާ@</font></div></td>
                </tr>
                <%
do while not rs.eof
%>
                <tr align="center"> 
                  <td width="16%" height="15"><%=rs("���~�W")%></td>
                  <td width="8%" height="15"> <%=rs("����")%> </td>
                  <td width="12%" height="15"><%=rs("����")%></td>
                  <td width="10%" height="15"> <%=rs("�ƶq")%> </td>
                  <td width="10%" height="15"><%=rs("����")%></td>
                  <td width="12%" height="15"><%=rs("���s")%></td>
                  <td width="12%" height="15"><a href="longzhuok.asp?id=<%=rs("ID")%>">��}</a></td>
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
            </div>
          </td>
          <td width="121"> 
            <div align="center"></div>
          </td>
        </tr>
      </table>
      <br>
      �s�]�i�b����̥��Ǫ��H����o�A���v��80%<br>
      �F�C�`���۳СI <br>
  </table>

</body>

</html>
