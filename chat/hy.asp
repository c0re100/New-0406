<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from �奻 "
Set Rs=conn.Execute(sql)
%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
A {COLOR: #000000; TEXT-DECORATION: none; TEXT-TRANSFORM: none}
A:hover {COLOR:#C46200; TEXT-DECORATION: underline}
BODY {FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt}
</style>
<title>�F�C�����`���|��²��</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3366CC" text="#FFFFFF">
<br>
<table border="0" width="375" cellspacing="1" cellpadding="1" align="center" bgcolor="#0099FF">
  <tr> 
    <td width="100%"> 
      <p align="center"><span style="FONT-SIZE: 14px"><b><font color="#FFFF00">�F�C����|��²���]2002.8.10)</font></b></span> 
    </td>
  </tr>
</table>
<br>
<table border="0" width="619" cellspacing="0" cellpadding="0" height="427" align="center" bgcolor="#3366CC">
  <tr> 
    <td width="1" height="897">&nbsp;</td>
    <td width="617" height="897"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>�@�B������n�}�l�|���\��H</b></font></td>
        </tr>
      </table>
      <p>�{�b�n���Ŷ�������A�n���j�a�@�Ӧn���T�����ҴN�����o���@�Ӧn���A�Ⱦ��A�{�b�A�Ⱦ����λ���ܩ��Q�A�g�L�F�h���V�O���M�S���ѨM�g�O�W�����D�A�O�S����k�Ӭ����I�|�����ާ@�O�b���v�T�Ҧ����a�����`�ϥΫe�D�U�i�檺�I<font color="#FFFFFF"><font color="#000000"><br>
        </font></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00"><font size="3">�G�B �|���ݹJ�H</font></font></td>
        </tr>
      </table>
      <p><font color="#FFFFFF"><font color="#000000"></font></font>�|������4�ءA����p�U�G<br>
        <br>
      </p>
      <table width="74%" border="1" cellspacing="0" cellpadding="2" align="center" height="74" bgcolor="#9966CC" bordercolor="#9966CC">
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FF99FF"><b><font color="#FFFFFF" size="2">1�ŷ|��</font><font color="#FFFFFF">�G</font></b></font><font color="#FFFF00"><b><%=rs("�|����I1")%>/��</b></font></font></font></font></font>�A�e�Ȩ�1000�U�A�[�԰�����<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFF00"><%=rs("�|������1")%>��</font></b></font></font></font></font>�A<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00">�w�I�ơG<%=rs("�|���w�I1")%>�I</font></font></font></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>2�ŷ|���G</b></font><font color="#FFFF00"><b><%=rs("�|����I2")%>��/��</b></font></font></font></font></font>�A�e�Ȩ�2000�U�A�[�԰�����<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFF00"><%=rs("�|������2")%>��</font></b><font color="#FFFFFF">�A<font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">�w�I�ơG<%=rs("�|���w�I2")%>�I</font></font></font></font></font></font></font></font></font></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"> 
              </font></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>3�ŷ|���G</b></font><font color="#FFFF00"><b><%=rs("�|����I3")%>��/��</b></font></font></font></font></font>�A�e�Ȩ�5000�U�A�[�԰�����<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("�|������3")%>��</b></font><font color="#FFFFFF">�A<font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">�w�I�ơG<%=rs("�|���w�I3")%>�I</font></font></font></font></font></font></font></font><font color="#000000"></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>4�ŷ|��</b><b>�G</b></font><font color="#FFFF00"><b><%=rs("�|����I4")%>��/��</b></font></font></font></font></font>�A�e�Ȩ�8000�U�A<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF">�[�԰�����<font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("�|������4")%>��</b></font></font></font></font></font></font></font></font>�A<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">�w�I�ơG<%=rs("�|���w�I4")%>�I</font></font></font></font></font></font></div>
          </td>
        </tr>
      </table>
      <p> <font size="3"> <font size="2" color="#FFFF00">�|���ʶR�ī~�B�˳Ʈɥ�5��B���}�������|�M�Ȩ�B���O���A�W�r�H�S�O�覡���</font></font></p>
      <p><font size="3"><font size="2" color="#FFFF00">1�ŷ|���G�|�O=500�� ����=����+500��<br>
        2�ŷ|���G�|�O=1000�� </font><font size="3"><font size="2" color="#FFFF00">����=����+1000��</font></font><font size="2" color="#FFFF00"><br>
        3�ŷ|���G�|�O=1500�� </font><font size="3"><font size="2" color="#FFFF00">����=����+1500��</font></font><font size="2" color="#FFFF00"><br>
        4�ŷ|���G�|�O=2000�� </font><font size="3"><font size="2" color="#FFFF00">����=����+2000��</font></font></font></p>
      <p><font size="3"><font size="2" color="#FFFF00"> �|�O�B�����i�H�ΨӶR<a href="../card/card.asp" target="_blank"><font color="#00FFFF">�|���d��</font></a>,�Q�����p��10%(����30%)�C</font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#99FF99" size="3"><br>
        <br>
        </font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF">�F�C����|������I�B�b�~�I�B�~�I�I<b><font color="#FFFF00"><br>
        <font size="+1"> </font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("�b�~�I")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font><b><font color="#FFFF00"></font></b></font></font></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>�T�B�涵�ʶR<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p>�԰��šG100��30��<br>
        �� ��G10��5000�U<br>
        �| �O�G10��20��</p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>�|�B�ר��|��<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p><font size="3"><font size="2" color="#00CCFF"><b><font color="#00FFFF">�ר��|��(3��):</font><font color="#FFFF00">150���A�e�԰�����40�šA�Ȩ�5000�U�A�w�I�ơG12�I 
        �]�`�G�ר��|�����ݩ�䥻�H�ϥΦ��ġA���o�~�ɩΰe�H�A�_�h�Ȱ���b���I�^ <br>
        ����@���A�H�ᤣ�ΦA��<br>
        �ר��|�����Ф@�Ӳר��|�����y�԰���10�šB�Ȩ�1000�U�B�|�O10��</font></b></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><br>
        <br>
        </font></font>�L�׬O��O�|���٬O�䥦���ޡB�x���A�p���H�Ϧ���W�w�A�N�ä[���o���ΡA�]�Q�}�����|���ڭ̤��h�^�ʶR�ɥ�Ǫ��O�ΡA�Фj�a�`�N�I</font><br>
        <font color="#FF0000"><b><font color="#FFFF00"><br>
        </font></b></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>���B�I�ڤ�k<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("�|���s�d")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></p>
      <p><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("�|���a�}")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></p>
    </td>
    <td width="10" height="897">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1" height="22">&nbsp;</td>
    <td width="617" height="22"> 
      <p align="center"><a href="javascript:window.close()"><font color="#FFFFFF">�������f</font></a> 
    </td>
    <td width="10" height="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1" height="15"></td>
    <td width="617" height="15"></td>
    <td width="10" height="15"></td>
  </tr>
  <tr> 
    <td width="1" bgcolor="#000099" height="2"></td>
    <td width="617" bgcolor="#000099" height="2"></td>
    <td width="10" bgcolor="#000099" height="2"></td>
  </tr>
</table>  
  
</body>  
  
</html>  
<%
	set rs=nothing	
	conn.close
	set conn=nothing %>