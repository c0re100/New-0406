<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp" -->
<%
dim copyright
dim Aff
Dim strSQL
Dim List,IDS,counter
IDS=Request("ID")
session("mynameup")=""
IF IDS="" Then
REsponse.Write "���~�I�ᥢ�H���I"
Response.End
End If
Aff = Chr(34)

'set list=server.createobject("adodb.recordset")
'list.Open "select * from list where id="+cint(Request.Querystring("IDS")) ,conn,1
'list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'�ƾڿ��~�B�z
'counter=list("counter")
'counter=counter + 1


'strSQL="SELECT * FROM List ID="+cint(Request.Querystring("IDS"))
'strSQL =strSQL & " Order by ID"
'set list=server.createobject("ADODB.RECORDSET")
'list.open strSQL,conn,3,2

strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  ID    LIKE '"&request("id")&"'"
strSQL =strSQL & " Order by ID DESC"

set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
if list.eof and list.bof then
Response.WRite "���~�I�S����e�Τ�I"
Response.End
End If

'if list.eof and list.bof then
' Response.WRite "���~�I�S����e�Τ�I"
' Response.End
'End if
counter=list("counter")
counter=counter+1
list("counter")=counter
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ӤH�H��--<%=List("name")%></title><STYLE type=text/css>BODY {
FONT-SIZE: 12px; FONT-FAMILY: "�s�ө���","Arial Narrow","Times New Roman"
}
P {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "�s�ө���","Arial Narrow","Times New Roman"
}
TD {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "�s�ө���","Arial Narrow","Times New Roman"
}
A:link {
FONT-SIZE: 12px; COLOR: #978de9; FONT-FAMILY: �s�ө���; TEXT-DECORATION: none
}
A:visited {
FONT-SIZE: 12px; COLOR: #978de9; TEXT-DECORATION: none
}
P {
FONT-SIZE: 12px; FONT-FAMILY: �s�ө���
}
DIV {
FONT-SIZE: 12px; FONT-FAMILY: �s�ө���
}
A:hover {
COLOR: red
}
</STYLE>
<script LANGUAGE="JavaScript">
bg = new Array(33);
bg[0] = '../riji/000.gif'; bg[1] = '../riji/001.gif'; bg[2] = '../riji/002.gif'; bg[3] = '../riji/003.gif'; bg[4] = '../riji/004.gif'; bg[5] = '../riji/005.gif'; bg[6] = '../riji/006.gif'; bg[7] = '../riji/007.gif'; bg[8] = '../riji/008.gif'; bg[9] = '../riji/009.gif'; bg[10] = '../riji/010.gif'; bg[11] = '../riji/011.gif'; bg[12] = '../riji/012.gif'; bg[13] = '../riji/013.gif'; bg[14] = '../riji/014.gif'; bg[15] = '../riji/015.gif'; bg[16] = '../riji/016.gif'; bg[17] = '../riji/017.gif'; bg[18] = '../riji/018.gif'; bg[19] = '../riji/019.gif'; bg[20] = '../riji/010.gif'; bg[21] = '../riji/021.gif'; bg[22] = '../riji/022.gif'; bg[23] = '../riji/023.gif'; bg[24] = '../riji/024.gif'; bg[25] = '../riji/025.gif'; bg[26] = '../riji/026.gif'; bg[27] = '../riji/027.gif'; bg[28] = '../riji/028.gif'; bg[29] = '../riji/029.gif'; bg[30] = '../riji/030.gif'; bg[31] = '../riji/031.gif'; bg[32] = '../riji/032.gif';
index = Math.floor(Math.random() * bg.length); document.write("<BODY BACKGROUND="+bg[index]+">"); </script> </head>

<body background="../ljimage/11.jpg">
<form method="POST" action="main.asp" >
<div align="center">
<input name=action type=hidden value=search>
<center>
<table border="0" width="600" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
<p align="center"><%=List("name")%>���`�U�H��</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" cellspacing="0" cellpadding="3" bordercolor="#988CD8" style="font-family: �s�ө���; font-size: 9pt" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="center"><a href=javascript:history.back();>��-��^�e��</a></td>
<td align="center"><a href=main.asp?action=search>��-�Ҧ��j�L</a></td>
<td align="center"><a href="add.asp">��-�j�L�[�J</a></td>
<td align="center"><a href="userre.asp"></a><a href="userre.asp">��-�ק���</a></td>
<td align="center"><a href="sf.asp?Keys=Login">��-�W�ź޲z</a></td>
<td align="center"><a href="admsearch.asp?keys=admsearch">��-���ŷj��</a></td>
<td align="center"><font color="#996600">�{��</font><font color="#FF6600">
<!--#include file="zongshu.asp" -->
</font><font color="#996600">��j�L�[�J</font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
      <table border="1" width="600" cellspacing="0" style="font-family: �s�ө���; font-size: 9pt" cellpadding="3" bordercolor="#988CD0" height="272">
        <tr> 
          <td bgcolor="#988CD0" colspan="6" height="12">�o�O<%=List("name")%>���n�����p</td>
        </tr>
        <tr> 
          <td  colspan="6" bgcolor="#EEF1F7" height="12" align=center> 
            <% if isnull(list("photo")) or isempty(list("photo")) then%>
            <a href style='cursor:hand;'onClick="window.open('iwantup.asp','_top','scrollbars=yes,resizable=yes,width=600,height=330')" title='�W�Ǭۤ�'>�٨S���W�Ǭۤ�</a> 
            <%else %>
            <img src="showimg.asp?id=<%=List("name")%>"> 
            <%End If%>
          </td>
        </tr>
        <tr> 
          <td width=74 align="right" bgcolor="#CCCCFF" height="12">�m�W�G</td>
          <td height="12" width="112"><font color=0000ff size=5 face=����,����><%=List("name")%></font></td>
          <td width=80 align="right" bgcolor="#CCCCFF" height="12">�����G</td>
          <td height="12" width="121"><%=list("mode")%></td>
          <td width=40 align="right" bgcolor="#CCCCFF" height="12">�ʧO�G</td>
          <td height="12" width="123"><%=list("sex")%></td>
        </tr>
        <tr> 
          <td align="right" bgcolor="#CCCCFF" height="11" width="74">�l�F�s�X�G</td>
          <td height="11" width="112"><%=list("pc")%></td>
          <td align="right" bgcolor="#CCCCFF" height="11" width="80">���W�١G</td>
          <td height="11" colspan="3"><%=list("danwei")%></td>
        </tr>
        <tr> 
          <td align="right" bgcolor="#CCCCFF" height="11" width="74">�ԲӦa�}�G</td>
          <td height="11" colspan="5"><%=list("address")%></td>
        </tr>
        <tr> 
          <td align="right" bgcolor="#CCCCFF" height="12" width="74">²�u�d���G</td>
          <td height="12" colspan="5"><%=list("doc")%></td>
        </tr>
        <tr> 
          <td  colspan="6" bgcolor="#CCCCFF" height="25"> 
            <p align="center"> 
              <input class="p9" type="submit" value="  ��^��ͦC��  ">
            </p>
          </td>
        </tr>
      </table>
</center>
</div>
</form>
</body>

</html>
<%List.CLose%>
<%Response.End%>
