<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();}</script>"
	Response.End 
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,���~�W,����,�ƶq,���O,��O,����,���s,�Ȩ� FROM ���~ WHERE �֦���="& info(9) & " and �ƶq>0 and �˳�=false and ����<>'�d��' order by ���� ",conn
%>
<html>
<head>
<title>���~�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#660000" background="../../chat/bk.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<div align="left">
  <div align="center"> </div>
  <div align="center">
    <table border="0" align="center" width="640" cellpadding="2" cellspacing="2" height="31">
      <tr align="center" bgcolor="#990000"> 
        <td width="53" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">���~�W</font></div></td>
        <td width="30" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">���� 
            </font></div>
        <td width="40" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">�ƶq 
            </font> </div>
        <td width="41" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">���O 
            </font></div>
        <td width="42" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">��O 
            </font></div>
        <td width="39" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">���� 
            </font></div>
        <td width="42" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">���s 
            </font></div>
        <td height="16" colspan="2" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">����</font></div></td>
        <td width="56" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">�覡</font></div></td>
        <td width="48" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">�ƶq</font></div></td>
        <td width="54" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">�ާ@</font></div></td>
      </tr>
      <%
do while not rs.eof
%>
      <tr bgcolor="#FFFFFF"> 
        <form method=POST action='fs.asp?id=<%=rs("id")%>&name=<%=name%>'>
          <td width="53" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("���~�W")%> </font> </div></td>
          <td width="30" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("����")%></font> </div>
          <td width="40" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("�ƶq")%> </font> </div>
          <td width="41" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("���O")%></font> </div>
          <td width="42" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("��O")%></font> </div>
          <td width="39" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("����")%></font> </div>
          <td width="42" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("���s")%></font> </div>
          <td colspan="2" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("�Ȩ�")%> </font> </div></td>
          <td height="3" width="56"> <div align="center"><font color="#FFFFFF"> 
              <select name="wpfs">
                <option value="1" selected>&nbsp;�ذe</option>
                <option value="2">&nbsp;���</option>
                <option value="3">&nbsp;�G��</option>
                <option value="4">�O�I�d</option>
              </select>
              </font></div></td>
          <td height="3" width="48"> <div align="center"> <font color="#FFFFFF"> 
              <select name="wpsl">
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
              </font></div></td>
          <td height="3" width="54"> <div align="center"><font color="#FFFFFF"> 
              <input type="SUBMIT" name="Submit"  value="�i��">
              </font></div></td>
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
    <table width="640" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><font color="#FFFFFF">�����G�b�o�ɧA�i�H��ۤv�����~�i��޲z�ާ@�C<br>
          1.�ذe�G��ۤv���F���ذe���o�]�L�^�A���ާ@���O�o�ͦb���ʤ����C�i�H�Φ��i���y�C<br>
          2.����G�A�i�H�P�A���B�Ͷi�檫�~����A�O�H�O�R���쪺�C�����ѧA�w�A�˥S�̩������C<br>
          3.�G��G�o�̧A�i�H���ѭn���A�@���@��A�@���@�R�C���L���H���{���F�j�Y�A�ڭ̥i���ޡC<br>
          4.�O�I�d�G������F��Q�ۤv�����s�_�ӡA�o�̥i�O�n�a��r�C�O���|�᪺�A���L���O��C</font></td>
</tr>
</table>
</div>
</div>
</body>
</html>
