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
if info(2)<10   then Response.Redirect "../error.asp?id=900"

wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ���~�R where id="&wupinid
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Redirect "modifywupin.htm"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="0064.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font size="+6" color="#FF0000">���~�ק�</font><font size="2"><br>
  <br>
  �`�G�����ק�ж��ī~�A�t���A�L���A�Y���A���ҡA���}�A�˹��A�r��,�d���E�������A�_�h�K�[�L�ġC</font></p>
<form method="post" action="modifybqnow.asp">
  <table border="1" cellspacing="1" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="305" bgcolor="#FFFFFF" cellpadding="1">
    <tr bgcolor="#99CCFF"> 
      <td width="105"><font size="2">ID</font></td>
      <td width="189" bgcolor="#99CCFF"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinid" readonly
value="<%=rst("ID")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">���~�W</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinname"
value="<%=rst("���~�W")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinlx"
value="<%=rst("����")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">���O</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinnl"
value="<%=rst("���O")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">��O</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupintl"
value="<%=rst("��O")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupingj"
value="<%=rst("����")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">���s</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinfy"
value="<%=rst("���s")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">��T��</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinjg"
value="<%=rst("��T��")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinsm"
value="<%=rst("����")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">����</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupindj"
value="<%=rst("����")%>" size="20">
        </font></td>
    </tr>
    <tr> 
      <td width="105"><font size="2">�Ȩ�</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinyl"
value="<%=rst("�Ȩ�")%>" size="20">
        </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="�T �w">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='javascript:history.go(-1)'" value="�� �^" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<div align="center"> 
  <p><font size="2">�ƾڮw��s�{�Ǧ]���ɶ������A�S���[�J���ŭȮɪ��˴��A�Фj�a���ɪ`�N�S���Ȫ��a���0</font> </p>
</div>
