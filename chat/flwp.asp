<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
fl=Trim(Request.QueryString("fl"))
if fl<>"�k��" and fl<>"����" and fl<>"����" and fl<>"�Y��" and fl<>"���}" and fl<>"�˹�" and fl<>"�ī~" and fl<>"�r��" and fl<>"�d��" then
	Response.Write "<script Language=Javascript>alert('�����d���ର�G�k��B����B���ҡB�Y���B���}�B�˹��B�ī~�B�r�ġB�d���A�Э��s��ܡI');window.close();</script>"
	Response.End
end if
if info(2)<7 then 
	Response.Write "<script Language=Javascript>alert('�A�Q�@����u�I');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM ���~ WHERE ����='" & fl & "' and �ƶq>0  order by �֦���",conn
%>
<html>
<head>
<title>���~�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center"><font color=yellow><b><%=fl%>�����~</b></font>�@��(�˳ƪ����~���~)<font face="����"><a href="javascript:this.location.reload()">��s</a></font></div>
<br>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">�֦���</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���~�W</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">�ƶq </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���O </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��O </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���s </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">����</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">�˳�</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">�覡</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�֦���")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���~�W")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("����")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�ƶq")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���O")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("��O")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("����")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���s")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�Ȩ�")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�˳�")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">�R��</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM ������� WHERE ����='" & fl & "' order by �֦���,�覡",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">�֦���</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���~�W</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">�ƶq </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���O </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��O </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���s </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">����</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">�覡</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">�ާ@</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�֦���")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���~�W")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("����")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�ƶq")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���O")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("��O")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("����")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("���s")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�Ȩ�")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("�覡")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">�R��</font></a></font></div>
      </td>
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
<tr align="center"> 
  <td nowrap width="45" height="16"> 
    <div align="center"></div>
  </td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td height="25"> 
      <div align="center">�o�̥i�H�d�ݨ��誺���~�A�R���N�i�H�N���Τ᪺���~�R���I</div>
    </td>
  </tr>
</table>
</body>
</html>
 