<!--#Include file="config.asp" -->
<%
'�]�m�K�X�M�b��

	On Error Resume Next

If Request("name1")="" or Request("password1")="" then
	Response.Write "�����L�k�T�{�A����B��I"
	Response.End
	Else
	Response.Write "���Ҩ���...<br>"
	If Request("name1")=username and Request("password1")=password then
		Response.Write "ok..�q�L���ҡI"
		sysop = True
		Session("sysop")= sysop
		call Croom
		Else
		Response.Write "Error!�L�k�q�L���ҡA���ˬd�ۤv�������I"
		sysop = False
		Session("sysop") = sysop
		response.End
	End If
Response.End
end if

Sub Croom()
	%>
<HTML>

<HEAD>
<META HTTP-EQUIV="Content-Language" CONTENT="zh-cn">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<META NAME="GENERATOR" CONTENT="Microsoft FrontPage 4.0">
<META NAME="ProgId" CONTENT="FrontPage.Editor.Document">
<TITLE>�t�κ޲z</TITLE>
<STYLE>
<!--
body         { font-size: 9pt }
-->
</STYLE>
</HEAD>

<BODY>

<P>�п�J�n�W�[���ж��W�١A�C�@�өж������Шϥ� , 
�r�����ζ}�ӡA�������s�b�Ů�A�Y�M�����s�b�r���A�Ф��n�ϥΤ��媺�r���C<BR>
�ܨҡG�ͱ���ѫ�,�˱���ѫ�,�R����ѫ�,�^���ѫ�,�����ѫ�,�t�β�ѫ�,�Y�Y���a</P>
<P><FONT COLOR="#FF0000">�`�N�G�p�G���ӳW�w�榡��J�A�N�o�ͨt�ο��~�I�I</FONT></P>
<FORM METHOD="POST" ACTION="sysop1.asp">
<input type=hidden name='action' value='sysop'>
<P><INPUT TYPE="text" NAME="cr" SIZE="100" STYLE="WIDHT:100% " VALUE="<%=OpenTextFile(Server.MapPath("."), "roomname.asp")%>"><BR>
<INPUT TYPE="submit" VALUE="����" NAME="B1"><INPUT TYPE="reset" VALUE="�������g" NAME="B2"></P>
</FORM>

</BODY>

</HTML>

	<%
End Sub

Function OpenTextFile(PathInfo, FileIn)
	On Error Resume Next
	Dim FileObject, InStream
	Set FileObject = Server.CreateObject("Scripting.FileSystemObject")	
	Set InStream = FileObject.OpenTextFile (PathInfo & "\" & FileIn, 1, False, False)	
	OpenTextFile = Instream.ReadALL
	InStream.Close
	Set InStream = Nothing
	If Err Then
		OpenTextFile = "���~��#" & Err.number & VbCrlf
		OpenTextFile = OpenTextFile & Err.description 
		Err.Clear
	End If
End Function
%>