<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "�жi�J��ѫǦA�i�J��Q�ާ@�I�I"
window.close()
</script>
<%response.end
end if%>
<html>
<head>
<title>��Q</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#660000" background="../../ff.jpg" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <table width="491" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="center"> <table width="491" align="center" cellspacing="0" border="1"
cellpadding="0">
          <tr> 
            <td bgcolor="#FFCC00"> 
              <div align="center"><font color="#FF6600"> 
                <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
rs.open "SELECT ���~�W,����,�ƶq,�Ȩ�,id FROM ���~ WHERE ����<>'�d��' and �˳�=false and �֦���=" & info(9) & "  order by �Ȩ� ",conn
if rs.bof and rs.eof then
%>
                ��Q�ѪO��s�D�G <br>
                �ѭ�</font>�V�V:-O<br>
                <%=name%>�A���W�@�L�Ҧ��A���|�O�Q��H�a�I<br>
                <%
else
%>
                <%=name%> �A���i���~���G </div></td>
          </tr>
        </table>
        <table border="0" cellspacing="2" cellpadding="2" width="491" align="center">
          <tr align="center" bgcolor="#CC0000"> 
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">�� 
              �~ �W</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">�� 
              �q</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">�� 
              ��</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">�{ 
              ��</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">��X�ƶq</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">�� 
              �@</font></td>
          </tr>
          <%
do while not rs.bof and not rs.eof
%>
          <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
            <form method=POST action='dan1.asp?id=<%=rs("id")%>'>
              <td><%=rs("���~�W")%></td>
              <td><%=rs("�ƶq")%></td>
              <td><%=rs("�Ȩ�")%>/��</td>
              <td><%=rs("�Ȩ�")/2%>/��</td>
              <td> <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select> </td>
              <td><div align="center"> 
                  <input type="SUBMIT" name="Submit" value="�i��">
                </div></td>
            </form>
          </tr>
          <%
rs.movenext
loop
end if
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