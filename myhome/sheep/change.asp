<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>更換寵物名字</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#3a4b91" background="../../ff.jpg" text="#000000" leftmargin="0" topmargin="0">
<p class="p1" align="center">&nbsp;</p>
<table border="1" align="center" width="281" cellpadding="0"
cellspacing="1" bordercolor="#000000" bgcolor="#0099CC">
<tr>
<td>
<form method="POST" action="change1.asp">
<table width="100%">
<tr>
<td colspan="2" height="31">
<div align="center">新名字：
<input type="text" name="name1" size="15">
</div>
</td>
</tr>
<tr>
<td align="center" colspan="2" height="2">
<input type="submit" value="修改"
name="submit">
<input type="reset" value="取消" name="reset">
</td>
</tr>
<tr>
<td align="center" colspan="2">&nbsp;</td>
</tr>
</table>
</form>
</td>
</tr>
</table>

</body>

</html>
