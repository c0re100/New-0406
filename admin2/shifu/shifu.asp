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
sql="select �v��,�Z�\�[,����,���O,��O,�ʧO,�m�W,grade from �Τ� where ID=" & info(9)
set rs=conn.execute(sql) 
shifu=rs("�v��")
wushu=rs("�Z�\�[")
zhizi=rs("����")
tili=rs("��O")
neili=rs("���O")%>
<html>

<head>
<title>�H�v�׷�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css"><!--td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { text-decoration: none; font-family: �s�ө���; font-size: 9pt }
a:hover      { text-decoration: underline; color: #CC0000; font-family: �s�ө���; font-size: 
               9pt }
.big         { font-family: �s�ө���; font-size: 12pt }
.txt         { font-family: �s�ө���; font-size: 10.8pt }
--></style>
</head>

<body BACKGROUND="../../ljimage/downbg.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#007cd0" align="center" width="389" cellpadding="8"
cellspacing="8"> <tr bgcolor="#FFFFFF"> <td height="210"> <table border="1"
      width="355" cellspacing="2" cellpadding="1" bordercolor="#65251C"> <tr> 
<td colspan="6" height="33"> <div align="center"> <font size="2" class="c"><font size="3"><b><FONT COLOR="#0099FF" SIZE="2"><IMG SRC='../../ico/<%=info(10)%>-2.gif' WIDTH='54' 

HEIGHT='54'></FONT><font color="#FF0033"><%=rs("�m�W")%></font></b><font
              size="2" color="#FF0033">�j�L������I��</font></font></font> </div></td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF"><font size="2" class="c">�m �W</font></td><td colspan="5" bgcolor="#FFFFFF"><%=rs("�m�W")%> 
����W���G <font color="#0000FF" size="2"> <%if rs("grade")<2  then response.write("��ӥE�D")               
if rs("grade")>=2and rs("grade")<3  then response.write("������")               
if rs("grade")>=3 and rs("grade")<4  then response.write("�p�����N")               
if rs("grade")>=4 and rs("grade")<5  then response.write("�n�W�㻮")               
if rs("grade")>=5 and rs("grade")<6  then response.write("��������")               
if rs("grade")>=6 and rs("grade")<7  then response.write("�@�N�j�L")               
if rs("grade")>=7 and rs("grade")<8  then response.write("�����C��")               
if rs("grade")>=8 and rs("grade")<9  then response.write("�D�W�ѤU")               
if rs("grade")>=9 and rs("grade")<10  then response.write("���C�P��")               
if rs("grade")>=10 then response.write("�C�P")               
            %> </font></td></tr> <tr> <td width="113" height="17" bgcolor="#FFFFFF"><font size="2" class="c">�� 
�O</font></td><td height="17" colspan="5" bgcolor="#FFFFFF"><%=rs("�ʧO")%> </td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF" height="10"><font size="2">���ɪ��W��</font><font size="2" class="c">:</font></td><td width="51" bgcolor="#FFFFFF" height="10"><%=rs("�Z�\�[")%> 
</td><td width="96" bgcolor="#FFFFFF" height="10"><font size="2">���G<%=rs("����")%></font></td><td width="67" bgcolor="#FFFFFF" height="10">�v�šG<font size="2"><%=rs("�v��")%></font></td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF" height="6">���O�G</td><td width="51" bgcolor="#FFFFFF" height="6"><%=rs("���O")%></td><td width="96" bgcolor="#FFFFFF" height="6">��O�G<%=rs("��O")%></td><td width="67" bgcolor="#FFFFFF" height="6"><a href="liangong.asp">��J�m�\��</a></td></tr> 
</table>�i�F�C���ܡj���v�Ū��j�L�i�H�b�o�M�A���v�žǲߪZ�\�A �ǲߤ@���A�ݭn������O100�A���20�A����<font color="#FF0000">�m�Z���W��</font><br> 
10�I�A�O�p�ݳo�m�Z�����@�C��}�A��<font color="#FF3333">�W��</font>���a���F <br> </table>

</body>

</html>
<%
conn.close
set conn=nothing
set rs=nothing%>
