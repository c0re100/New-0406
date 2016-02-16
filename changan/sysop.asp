<!--#Include file="config.asp" -->
<%
'設置密碼和帳號

	On Error Resume Next

If Request("name1")="" or Request("password1")="" then
	Response.Write "身份無法確認，停止運行！"
	Response.End
	Else
	Response.Write "驗證身份...<br>"
	If Request("name1")=username and Request("password1")=password then
		Response.Write "ok..通過驗證！"
		sysop = True
		Session("sysop")= sysop
		call Croom
		Else
		Response.Write "Error!無法通過驗證，請檢查自己的身份！"
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
<TITLE>系統管理</TITLE>
<STYLE>
<!--
body         { font-size: 9pt }
-->
</STYLE>
</HEAD>

<BODY>

<P>請輸入要增加的房間名稱，每一個房間之間請使用 , 
逗號分割開來，之間不存在空格，頭和尾不存在逗號，請不要使用中文的逗號。<BR>
示例：友情聊天室,親情聊天室,愛情聊天室,英文聊天室,中文聊天室,系統聊天室,某某的家</P>
<P><FONT COLOR="#FF0000">注意：如果不照規定格式輸入，將發生系統錯誤！！</FONT></P>
<FORM METHOD="POST" ACTION="sysop1.asp">
<input type=hidden name='action' value='sysop'>
<P><INPUT TYPE="text" NAME="cr" SIZE="100" STYLE="WIDHT:100% " VALUE="<%=OpenTextFile(Server.MapPath("."), "roomname.asp")%>"><BR>
<INPUT TYPE="submit" VALUE="提交" NAME="B1"><INPUT TYPE="reset" VALUE="全部重寫" NAME="B2"></P>
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
		OpenTextFile = "錯誤號#" & Err.number & VbCrlf
		OpenTextFile = OpenTextFile & Err.description 
		Err.Clear
	End If
End Function
%>