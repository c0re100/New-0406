<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>�D�㩱</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> <td align="center" valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr bgcolor="#FFFF00"> 
          <td width="17%" nowrap> 
            <div align="center"><a href="BINQI.asp"><font color="#6666FF">����M�C</font></a> 
</div></td>
          <td width="16%" nowrap> 
            <div align="center"><a href="dunpai.asp"><font color="#6666FF">����@��</font></a></div></td>
          <td width="16%" nowrap> 
            <div align="center"><a href="body.asp"><font color="#6666FF">����</font></a></div></td>
          <td width="12%" nowrap> 
            <div align="center"><a href="head.asp"><font color="#6666FF">�Y��</font></a></div></td>
          <td width="13%" nowrap> 
            <div align="center"><a href="foot.asp"><font color="#6666FF">���}</font></a> </div></td>
          <td width="11%" nowrap> 
            <div align="center"><a href="other.asp"><font color="#6666FF">�˹� </font></a></div></td>
          <td width="15%" nowrap> 
            <div align="center"><a
href="anqi.asp"><font color="#6666FF">�t��</font></a></div></td></tr> </table>
      <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#99CCFF">
        <tr align="center" bgcolor="#FF0000"> 
          <td width="64"><font color="#FFFFFF">�˳ƦW</font></td>
          <td width="85"><font color="#FFFFFF"> �Z �� ��</font></td>
          <td width="53"><font color="#FFFFFF">���� </font> 
          <td width="270"><font color="#FFFFFF">�\��</font></td>
          <td width="117"><font color="#FFFFFF">���</font></td>
          <td colspan="2"><font color="#FFFFFF">�ƶq</font></td>
          <td width="46"><font color="#FFFFFF">�ާ@</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT  ���~�W,id,����,����,���s,����,��T��,�Ȩ� FROM ���~�R where ����='�˹�' order by �Ȩ� ",conn
do while not rs.bof and not rs.eof
%>
        <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
          <form method=POST action='binqi1.asp?id=<%=rs("id")%>'>
            <td width="64"> <div align="center"><%=rs("���~�W")%></div></td>
            <td width="85"> <div align="center"><img src="001/<%=rs("����")%>.gif" border="0" alt="<%=rs("���~�W")%> " width="60" height="58"></div></td>
            <td width="53"> <div align="center"><%=rs("����")%></div></td>
            <td width="270"> <div align="center">���ŻݨD<%=abs(rs("����"))%>�A�[���s<%=abs(rs("���s"))%>�A��T��<%=abs(rs("��T��"))%></div></td>
            <td width="117"> <div align="right"><%=rs("�Ȩ�")%>��</div></td>
            <td colspan="2"> <div align="center"> 
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
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="46"> <div align="center"> 
                <input type="SUBMIT" name="Submit"
value="�i��">
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
