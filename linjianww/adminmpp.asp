<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>�����޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<script language="JavaScript">
function shutwin()
{
window.close();
return;
}
</script>
</head>
<body background="0064.jpg" text="#000000">
<p align="center"> <font size="+6"> 
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
set subj=server.createobject("adodb.recordset")
set owner=server.createobject("adodb.recordset")
subj.open "Select * From ����",conn,0,1
if not subj.eof then
%>
  <font color="#FF3333" face="����">[ �����޲z ]</font></font></p>
<p align="center"><font size="2">�b�����޲z�̭��A���ѥi�H���x���W�A�R���x�������A�M�ק�x����ơI</font></p>
<table width="505" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
  <tr bgcolor="#FF6600"> 
    <td align="center" width="55"><font color="#FFFFFF" size="2">&nbsp;�� ��</font></td>
    <td width="39" align="center"><font size="2" color="#FFFFFF">�̤l��</font></td>
    <td width="59" align="center"><font color="#FFFFFF" size="2">�ۦ��ʽ�</font></td>
    <td width="60" align="center"><font color="#FFFFFF" size="2">�x ��</font></td>
    <td width="103" align="center"><font size="2" color="#FFFFFF">�������</font></td>
    <td width="64" align="center"><font size="2" color="#FFFFFF">�}���x��</font></td>
    <td align="center" colspan="2"><font color="#FFFFFF" size="2">�� �@</font></td>
  </tr>
  <%do while not subj.eof%>
  <tr bgcolor="#FFFFFF"> 
    <td width="55"> 
      <div align="center"><font size="2"><%=subj("����")%></font></div>
    </td>
    <td width="39"> 
      <div align="right"><font size="2"><%=subj("�̤l��")%>��</font></div>
    </td>
    <td width="59"> 
      <div align="center"><font size="2"><%=subj("�A�X")%></font></div>
    </td>
    <td width="60"><font size="2"><%=subj("�x��")%></font></td>
    <td width="103"> 
      <div align="right"><font size="2"><%=subj("�������")%></font></div>
    </td>
    <td width="64"> 
      <div align="center"><font size="2"><a href=delzm.asp?username=<%=subj("����")%> title="�R���x��"> 
        </a><a href=delzm.asp?username=<%=subj("����")%> title="�R���x��">�}��</a></font></div>
    </td>
    <td width="32"> 
      <p align="center"><font size="2"><a href="delmp.asp?owner=<%=SUBJ("����")%>&action=delete" title="�R������">�R��</a></font></p>
    </td>
    <td width="35"> 
      <p align="center"><font size="2"><a href="managemp.asp?id=<%=SUBJ("����")%>">�ק�</a></font> 
    </td>
  </tr>
  <%subj.movenext
loop %>
  <% else %>
</table>
<table width="54%" border="0" cellspacing="0" cellpadding="0" height="40" align="center">
<tr>
<td>
      <div align="center"><font size="2">�R�������|�N�ɮפ��Ҧ��Ӫ������H--�����m�����L���A���������L���A�P�ɧR���׾¤��Ӫ������Ҧ����l�A���ѡA�Ҽ{�n�A�����R�����ڡC<br>
        (�R�������Ȯ��٤����x���H�����󪬺A�A���ѭn�[�W�]���)</font></div>
</td>
</tr>
</table>
<p>&nbsp;</p>
<table align="center" width="197">
<td height="14">
      <div align="center"><font size="2">�٨S���@�Ӫ����O�I</font> 
        <% subj.close
set subj=nothing
end if%>
      </div>
</table>
<div align="center"></div>
<p align="center"><font size="2" color="#FF0000">[<a href="adminmpp.asp">====��s====</a>]</font><br>
</body>
