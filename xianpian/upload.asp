<%
if Session("mynameup")="" then
%>
<script language="vbscript">
MsgBox "���n�W�Ǭۤ��A�бq[�ק���]�B�i�J�I�I"
location.href = "userre.asp"
</script>
<%
end if
%>

<!--#include file="conn.asp"-->
<% rs=conn.execute("select * from List where name='"&session("myname")&"'")%>
<HTML><HEAD><TITLE><%=session("myname")%>_�Ӥ��W��</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<style type="text/css">
<!--td {  font-family: �s�ө���; font-size: 9pt}
body {  font-family: �s�ө���; font-size: 9pt}
select {  font-family: �s�ө���; font-size: 9pt}
A {text-decoration: none; font-family: "�s�ө���"; font-size: 9pt}
A:hover {text-decoration: underline; color: #666666; font-family: "�s�ө���"; font-size: 9pt} .txt {  font-family: "�s�ө���"; font-size: 10.5pt}
--></style>

</HEAD>
<body bgcolor="#8d8051" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" background="../ljimage/11.jpg">
<a href=javascript:window.close();>��������</a> | <a href=javascript:location.reload();>��s���</a> 
<TABLE width=100% align="center">
<TBODY>
<TR>
<TD align=middle width="704" height="294">
<div align="right"><BR>
</div>
<P align="center"><%=session("myname")%>_�Ӥ��W��</a>
<form name="Form" enctype="multipart/form-data" action="upphoto.asp" method=post>
<TABLE border=1 cellPadding=0 cellSpacing=1 width="443" bordercolor="#000000" align="center">
<TBODY>
<tr>
<TD bgColor=#b2a265 vAlign=center>
<P>�b�A���w�L�W�n�W�Ǫ��Ϥ����: <INPUT name=mefile type=file>
<P align="center">
<input type="submit" name="Submit" value="�W�ǷӤ�">
<input type="reset" name="cancel" value="�u�r���F">
<br>
�`�N�G�����\�W�Ǥ��榡<font color="#00FFFF">.jpg</font>��<font color="#00FFFF">.gif</font>���Ӥ�</P>
</TD>
</TR>
<TR>
            <TD background="../ljimage/11.jpg">&nbsp;<br><center>
              <% if isnull(rs("photo")) or isempty(rs("photo")) then%>
              <img src="photo.jpg" width="160" height="221"> 
              <%else %>
              <img src="showimg.asp?id=<%=session("myname")%>"> 
              <%end if%>
              <br>
<font color="#000000"> �o�O�A���Ӥ���?</font> </TD></tr>          </TBODY>
</TABLE>
</FORM>
</TD>
</TR>
</TBODY>
</TABLE>
<center><a href=javascript:window.close();>��������</a> | <a href=javascript:location.reload();>��s���</a>
</BODY></HTML>
