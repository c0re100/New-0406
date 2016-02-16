<%
if Session("mynameup")="" then
%>
<script language="vbscript">
MsgBox "◎要上傳相片，請從[修改資料]處進入！！"
location.href = "userre.asp"
</script>
<%
end if
%>

<!--#include file="conn.asp"-->
<% rs=conn.execute("select * from List where name='"&session("myname")&"'")%>
<HTML><HEAD><TITLE><%=session("myname")%>_照片上載</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<style type="text/css">
<!--td {  font-family: 新細明體; font-size: 9pt}
body {  font-family: 新細明體; font-size: 9pt}
select {  font-family: 新細明體; font-size: 9pt}
A {text-decoration: none; font-family: "新細明體"; font-size: 9pt}
A:hover {text-decoration: underline; color: #666666; font-family: "新細明體"; font-size: 9pt} .txt {  font-family: "新細明體"; font-size: 10.5pt}
--></style>

</HEAD>
<body bgcolor="#8d8051" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" background="../ljimage/11.jpg">
<a href=javascript:window.close();>關閉本窗</a> | <a href=javascript:location.reload();>刷新顯示</a> 
<TABLE width=100% align="center">
<TBODY>
<TR>
<TD align=middle width="704" height="294">
<div align="right"><BR>
</div>
<P align="center"><%=session("myname")%>_照片上載</a>
<form name="Form" enctype="multipart/form-data" action="upphoto.asp" method=post>
<TABLE border=1 cellPadding=0 cellSpacing=1 width="443" bordercolor="#000000" align="center">
<TBODY>
<tr>
<TD bgColor=#b2a265 vAlign=center>
<P>在你的硬盤上要上傳的圖片文件: <INPUT name=mefile type=file>
<P align="center">
<input type="submit" name="Submit" value="上傳照片">
<input type="reset" name="cancel" value="哎呀錯了">
<br>
注意：隻允許上傳文件格式<font color="#00FFFF">.jpg</font>或<font color="#00FFFF">.gif</font>的照片</P>
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
<font color="#000000"> 這是你的照片嗎?</font> </TD></tr>          </TBODY>
</TABLE>
</FORM>
</TD>
</TR>
</TBODY>
</TABLE>
<center><a href=javascript:window.close();>關閉本窗</a> | <a href=javascript:location.reload();>刷新顯示</a>
</BODY></HTML>
