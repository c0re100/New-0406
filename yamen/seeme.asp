<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="Select  �Z�\�[,���O�[,�����[,���s�[,��O�[,�y�O�[,�D�w�[,�k�O�I,�m�W,���A,�ʧO,�|���O,����,��O,����,����,���O,�s��,�Z�\,�Ȩ�,����,�t��,���s,�G�B,�D�w,���ФH,�y�O,grade from �Τ� where �m�W='"&info(0)&"'"
set rs=conn.Execute(sql)
wgj=rs("�Z�\�[")
nlj=rs("���O�[")
gjj=rs("�����[")
fyj=rs("���s�[")
tlj=rs("��O�[")
mlj=rs("�y�O�[")
ddj=rs("�D�w�[")
flj=rs("�k�O�I")
%>
<html>

<head>
<title>�����ɮ�</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body topmargin="0" leftmargin="0" background="../images/8.jpg" text="#993333" vlink="#663366" link="#660066">
<table width="412" border="0" cellspacing="0" cellpadding="0" height="310" align="center">
  <tr>
    <td width="412" height="308"> 
      <table border="1"
width="407" cellspacing="1" cellpadding="1" align="center">
        <tr> 
          <td colspan="5" height="23"> 
            <table border="0"
width="400" cellspacing="1" cellpadding="1" align="center" background="../ljimage/downbg.jpg">
              <tr> 
                <td width="65" height="25"><font color="#0099FF" size="2"><img src='../ico/<%=info(10)%>-2.gif' width='54' 

height='54'></font></td>
                <td width="111" height="25"><font color="#0099FF" size="2"><%=rs("�m�W")%></font></td>
                <td width="63" height="25">�� �A�G</td>
                <td height="25" colspan="2" width="148"> <font color="#0099FF" size="2"><%=rs("���A")%></font></td>
              </tr>
              <tr> 
                <td width="65" height="25">�� �O�G</td>
                <td width="111" height="25"><font color="#0099FF" size="2"><%=rs("�ʧO")%></font></td>
                <td width="63" height="25">�| �O�G</td>
                <td height="25" colspan="2" width="148"> <font color="#0099FF" size="2"><%=rs("�|���O")%></font><font size="2">��</font></td>
              </tr>
              <tr> 
                <td width="65" height="25">�� ���G</td>
                <td width="111" height="25"><%=rs("����")%></td>
                <td width="63" height="25">�� �O�G</td>
                <td height="25" colspan="2" width="148"><%=rs("��O")%> / <font color="#0099FF" size="2"><%=rs("����")%></font></td>
              </tr>
              <tr> 
                <td width="65" height="20">�� ���G</td>
                <td width="111" height="24"><%=rs("����")%></td>
                <td width="63" height="24">�� �O�G</td>
                <td height="24" colspan="2" width="148"><%=rs("���O")%> / <font color="#0099FF" size="2"><%=rs("����")*640+2000+nlj%></font> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">�s �ڡG</td>
                <td width="111" height="24"><%=rs("�s��")%> ��</td>
                <td width="63" height="24">�Z �\�G</td>
                <td height="24" colspan="2" width="148"><%=rs("�Z�\")%> / <font color="#0099FF" size="2"><%=rs("����")*1280+3800+wgj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">�{ ���G</td>
                <td width="111" height="24"><%=rs("�Ȩ�")%> ��</td>
                <td width="63" height="24">�� ���G</td>
                <td height="24" colspan="2" width="148"><%=rs("����")%> / <font color="#0099FF" size="2"><%=rs("����")*1200+4500+gjj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">�t ���G</td>
                <td width="111"><%=rs("�t��")%> </td>
                <td width="63">�� �s�G</td>
                <td colspan="2" width="148"><%=rs("���s")%> / <font color="#0099FF" size="2"><%=rs("����")*1120+3000+fyj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">�G �B�G</td>
                <td width="111"><%=rs("�G�B")%></td>
                <td width="63">�D �w�G</td>
                <td colspan="2" width="148"><%=rs("�D�w")%><font size="2"> /</font> 
                  <font color="#0099FF" size="2"><%=rs("����")*1440+2200+ddj%></font> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">���ФH�G</td>
                <td width="111"><%=rs("���ФH")%> </td>
                <td width="63">�y �O�G</td>
                <td colspan="2" width="148"><%=rs("�y�O")%> <%=rs("����")*1120+2100+mlj%> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">������šG</td>
                <td width="111">�䯫</td>
                <td width="63">�԰��šG</td>
                <td colspan="2" width="148"><%=rs("����")%> ��</td>
              </tr>
              <tr> 
                <td width="65" height="20">Ĺ / ��G</td>
                <td width="111"> </td>
                <td width="63">�޲z�šG</td>
                <td colspan="2" width="148"><%=rs("grade")%> ��</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
      </td>
</tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
%> 