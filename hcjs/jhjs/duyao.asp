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
<title>�����ľQ</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<div align="center"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td align="center" valign="top"> <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
          <tr> 
            <td colspan="6" bgcolor="#FFFF00"> <div align="center"> 
                <div align="center"></div>
                <div align="center">[ <font color="#FF3300"><a href="yaopu0.asp"><font color="#6633FF">�� 
                  �� �M �� ��</font></a></font> ]--[ <font color="#FF0000">�r �� �M �� 
                  ��</font> ]</div>
              </div></td>
          </tr>
          <tr bgcolor="#FF0000"> 
            <td width="125"> <div align="center"><font color="#FFFFFF">�r ��</font></div></td>
            <td width="115" height="13"> <div align="center"><font color="#FFFFFF"> 
                �� �~ ��</font></div></td>
            <td width="244"> <div align="center"><font color="#FFFFFF">�\ ��</font></div></td>
            <td width="76"> <div align="center"><font color="#FFFFFF">�� ��</font></div></td>
            <td width="93"> <div align="center"><font color="#FFFFFF">�� �q</font></div></td>
            <td width="104"> <div align="center"><font color="#FFFFFF">�� �@</font></div></td>
          </tr>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,���~�W,���O,��O,�Ȩ�,���� FROM ���~�R where  ����='�r��' order by �Ȩ�",conn
do while not rs.bof and not rs.eof
%>
          <tr  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
            <form method=POST action='buyduyao.asp?id=<%=rs("id")%>'>
              <td width="125" bgcolor="#FFFFFF"> <div align="center"><%=rs("���~�W")%></div></td>
              <td width="115" bgcolor="#FFFFFF"> <div align="center"><img src="001/<%=rs("����")%>.gif" border="0" alt="<%=rs("���~�W")%> "></div></td>
              <td width="244" bgcolor="#FFFFFF">�����O<%=rs("���O")%>�A���ͩR<%=rs("��O")%></td>
              <td width="76" bgcolor="#FFFFFF"> <div align="center"><%=rs("�Ȩ�")%>��</div></td>
              <td width="93" bgcolor="#FFFFFF"> <div align="center"> 
                  <select name="sl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
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
              <td width="104" bgcolor="#FFFFFF"> <div align="center"> 
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
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
