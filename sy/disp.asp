<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim conn,rs,id
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
id=request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open ("SELECT * FROM �ӭ� WHERE ID=" & id),conn,0,1
%>

<html>
<head>
<title>��  -&gt; �d�ݪ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css><!--td {  font-family: �s�ө���; font-size: 9pt}body {  font-family: �s�ө���; font-size: 9pt}select {  font-family: �s�ө���; font-size: 9pt}A {text-decoration: none; font-family: "�s�ө���"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "�s�ө���"; font-size: 9pt} .big {  font-family: �s�ө���; font-size: 12pt}.txt {  font-family: "�s�ө���"; font-size: 10.8pt}
--></style>
</head>
<body bgcolor="#3a4b91" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<hr size="1">
<table width="590" border="0" cellspacing="0" cellpadding="3" align="center">
<tr>
<td>���`�H<b>:</b><b><%=rs("��i")%></b>&nbsp;�d�G�a�g�D�G</td>
<td align="right"><a href=javascript:history.back()>�� �^</a></td>
</tr>
</table>
<table width="588" border="1" cellspacing="1" cellpadding="0" align="center">
<tr >
<td height="67">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<p><span class="txt"><br>
<%=changechr(rs("����"))%><br>
<br>
<br>
��Q�i�o�شc�H��|�A�Ʊ浹��<%=rs("�n�D")%> <br>
<br>
</span></p>
<p align="right">&nbsp;</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%rs.close
set rs=nothing
%>
<hr size="1">
<p align="center">&nbsp;</p>
</body>
</html>
<%
function getorder(theid)
dim tmprs
tmprs=conn.execute("select [Order] from �ӭ� Where ID=" & theid)
getorder=tmprs("Order")
set tmprs=nothing
end function

function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%>
