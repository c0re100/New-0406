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
<title>�ī~��</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> <td align="center" valign="top"> 
      <table border="0" cellpadding="2" cellspacing="2" width="100%" bgcolor="#99CCFF">
        <tr bgcolor="#FFFF00"> 
          <td height="13" colspan="6"> 
            <div align="center"></div>
            <div align="center">[ <font color="#FF3300">�� �� �M �� ��</font> ]--[ 
              <a
href="duyao.asp"><font color="#6633FF">�r �� �M �� ��</font></a> ]</div></td>
        </tr>
        <tr> 
          <td width="124" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">�� 
              ��</font></div></td>
          <td width="99" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF"> 
              �� �~ ��</font></div></td>
          <td width="266" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">�\ 
              ��</font></div></td>
          <td width="91" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">�� 
              ��</font></div></td>
          <td width="85" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">�� 
              �q</font></div></td>
          <td width="92" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">�� 
              �@</font></div></td>
        </tr>
        <%
rs.open "SELECT id,���~�W,����,���O,��O,�Ȩ� FROM ���~�R where  ����='�ī~' order by �Ȩ�",conn
do while not rs.bof and not rs.eof
%>
        <tr  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyyao.asp?id=<%=rs("id")%>'>
            <td width="124" bgcolor="#FFFFFF"> 
              <div align="center"><%=rs("���~�W")%></div></td>
            <td width="99" bgcolor="#FFFFFF"> 
              <div align="center"><img src="001/<%=rs("����")%>.gif" border="0" alt="<%=rs("���~�W")%> "></div></td>
            <td width="266" bgcolor="#FFFFFF">�ɤ��O<%=rs("���O")%>�A�ɥͩR<%=rs("��O")%>!</td>
            <td width="91" bgcolor="#FFFFFF"> 
              <div align="center"><%=rs("�Ȩ�")%>��</div></td>
            <td width="85" bgcolor="#FFFFFF"> 
              <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="99">99</option>
                </select>
              </div></td>
            <td width="92" bgcolor="#FFFFFF"> 
              <div align="center"> 
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
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
    </td></tr> </table><br> 
</body>

</html>
